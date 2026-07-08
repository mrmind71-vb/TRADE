VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form VsShip_Pay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„ «»⁄… »Ê‰«  «·»Ì⁄  ”·Ì„ «Ê‰ ·«Ì‰"
   ClientHeight    =   11040
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   20400
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   20400
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   1140
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   45
      Width           =   7935
      Begin VB.Frame Frame4 
         Height          =   600
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   495
         Width           =   3345
         Begin VB.OptionButton XONLINE 
            Alignment       =   1  'Right Justify
            Caption         =   "‰ﬁœÏ"
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
            Index           =   0
            Left            =   2205
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   180
            Width           =   1005
         End
         Begin VB.OptionButton XONLINE 
            Alignment       =   1  'Right Justify
            Caption         =   "‘Õ‰"
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
            Index           =   1
            Left            =   1035
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   180
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.OptionButton XONLINE 
            Alignment       =   1  'Right Justify
            Caption         =   "«·ﬂ·"
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
            Index           =   2
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   180
            Width           =   825
         End
      End
      Begin VB.OptionButton XPAY 
         Alignment       =   1  'Right Justify
         Caption         =   "«·ﬂ·"
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
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   225
         Width           =   645
      End
      Begin VB.OptionButton XPAY 
         Alignment       =   1  'Right Justify
         Caption         =   "„Õ’·"
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
         Index           =   1
         Left            =   1095
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   225
         Width           =   870
      End
      Begin VB.OptionButton XPAY 
         Alignment       =   1  'Right Justify
         Caption         =   "€Ì— „Õ’·"
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
         Index           =   0
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   225
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.TextBox xDatePAY2 
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
         Left            =   3465
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   630
         Width           =   1410
      End
      Begin VB.TextBox xDatePAY1 
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
         Left            =   4995
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   630
         Width           =   1410
      End
      Begin VB.TextBox xDate2 
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
         Left            =   3465
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   225
         Width           =   1410
      End
      Begin VB.TextBox xDate1 
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
         Left            =   4995
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   225
         Width           =   1410
      End
      Begin VB.Label Label2 
         Caption         =   " «—ÌŒ  Õ’Ì· : "
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
         Index           =   2
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   675
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   " «—ÌŒ „»Ì⁄«   : "
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
         Index           =   0
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   270
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   45
      Width           =   12075
      Begin VB.TextBox XPHONE 
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
         Left            =   8325
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   990
         Width           =   2220
      End
      Begin VB.TextBox XONLINE_DOC 
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
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1035
         Width           =   1635
      End
      Begin VB.TextBox XSHIP_NO 
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
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   630
         Width           =   1635
      End
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2520
         Picture         =   "VsShip_Pay.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   90
         Picture         =   "VsShip_Pay.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   1305
         Picture         =   "VsShip_Pay.frx":495E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   135
         Width           =   1185
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   7110
         TabIndex        =   4
         Top             =   180
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   90
         TabIndex        =   9
         Top             =   630
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
         Picture         =   "VsShip_Pay.frx":6D88
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   330
         Left            =   3825
         TabIndex        =   10
         Top             =   180
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   582
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSDataListLib.DataCombo xship 
         Height          =   315
         Left            =   7110
         TabIndex        =   25
         Top             =   630
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "—ﬁ„ «· ·Ì›Ê‰ : "
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
         Left            =   10620
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1035
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "—ﬁ„ «·ÿ·»Ì… : "
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
         Index           =   5
         Left            =   5805
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "—ﬁ„ «·»Ê»Ì’… : "
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
         Index           =   4
         Left            =   5805
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   675
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "‘—ﬂ… «·‘Õ‰ : "
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
         Index           =   3
         Left            =   10575
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   675
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "«·›—⁄ :"
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
         Index           =   1
         Left            =   10620
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   8925
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1755
      Width           =   20040
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   8610
         Left            =   135
         TabIndex        =   2
         Top             =   180
         Width           =   19815
         _cx             =   34951
         _cy             =   15187
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10710
      Width           =   20400
      _ExtentX        =   35983
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   17639
            MinWidth        =   17639
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc DATA1 
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
   Begin MSAdodcLib.Adodc DATA2 
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
   Begin MSAdodcLib.Adodc data3 
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
Attribute VB_Name = "VsShip_Pay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    cHead1 = Me.Caption & "  " & XSTORE.text
    grid1.ColHidden(15) = True
    
    grid1.ColHidden(0) = (XSTORE.BoundText <> "")
    
    Set printGrd3.grid1 = grid1
    Set printGrd3.myForm = Me
    If XSTORE.BoundText = "" Then
        printGrd3.doprint 1, 0, cHead1, , , False, True, 10, "Arial", 0
    Else
        printGrd3.doprint 0.85, 0, cHead1, , , False, True, 10, "Arial"
    End If
    printGrd3.Show 1
    grid1.ColHidden(15) = False
    grid1.ColHidden(0) = False
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
    openCon con
    
    Set DATA2.Recordset = myRecordSet("Select Code,DescA From STORE_BR ORDER BY CODE ", con)
    Set XSTORE.RowSource = DATA2
    XSTORE.ListField = "Desca"
    XSTORE.BoundColumn = "Code"
    If cBranch <> "00" Then XSTORE.BoundText = cBranchStore
    XSTORE.Enabled = (cBranch = "00")
    
    Set data3.Recordset = myRecordSet("Select Code,DescA From ship ORDER BY CODE ", con)
    Set xship.RowSource = data3
    xship.ListField = "Desca"
    xship.BoundColumn = "Code"
    
    Set grid1.DataSource = DATA1
    grid1.Rows = 1
    Fixgrd
End Sub
Private Sub myload()
    Dim cWhere As String, cString As String
   '                        0       1       2       3       4       5       6       7                       8       9       10      11      12          13          14          15          16
    cString = " SELECT  STORENAME, date, DOC_NO2, NAME, phone, CHARGE1, CHARGE2, TOTAL, (TOTAL+CHARGE1+CHARGE2 ), DESCA, SHIP_NO, BOXNAME , BANKNAME , DATE_PAY , DOC_NO , ' ', note_ret , ONLINE_DOC FROM Q_SHIP WHERE DOC_NO IS NOT NULL AND ONLINE > 0 "
    If XSTORE.BoundText <> "" Then cString = cString & "  AND store = " & MyParn(XSTORE.BoundText)
    If xship.BoundText <> "" Then cString = cString & "  AND ship = " & MyParn(xship.BoundText)
    If XSHIP_NO.text <> "" Then cString = cString & "  AND ship_NO = " & MyParn(XSHIP_NO.text)
    
    If XPHONE.text <> "" Then cString = cString & "  AND phone = " & MyParn(XPHONE.text)
    If XONLINE_DOC.text <> "" Then cString = cString & "  AND ONLINE_DOC = " & MyParn(XONLINE_DOC.text)
    
    If IsDate(xDate1.text) Then cString = cString & " AND DATE >= " & DateSq(xDate1.text)
    If IsDate(xDate2.text) Then cString = cString & " AND DATE <= " & DateSq(xDate2.text)
    
    If IsDate(xDatePAY1.text) Then cString = cString & " AND DATE_PAY >= " & DateSq(xDatePAY1.text)
    If IsDate(xDatePAY2.text) Then cString = cString & " AND DATE_PAY <= " & DateSq(xDatePAY2.text)
    
    If XPAY(1).Value Then cString = cString & " AND (DATE_PAY IS NOT NULL OR ONLINE =   1 )"
    If XPAY(0).Value Then cString = cString & " AND (DATE_PAY IS NULL AND ONLINE =  2 )"
    
    If XONLINE(0).Value Then cString = cString & " AND ONLINE = 1 "
    If XONLINE(1).Value Then cString = cString & " AND ONLINE = 2 "
    
    cString = cString & " order by STORE , DATE , doc_no2 "
    Set DATA1.Recordset = myRecordSet(cString, con)
    Fixgrd
End Sub
Sub Fixgrd()
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .WordWrap = True
    .Cols = 18
    .RowHeight(0) = 1000
    .TextMatrix(0, 0) = "«·›—⁄"
    .TextMatrix(0, 1) = " «—ÌŒ"
    .TextMatrix(0, 2) = "—ﬁ„ "
    .TextMatrix(0, 3) = "«·⁄„Ì·"
    .TextMatrix(0, 4) = "«· ·Ì›Ê‰"
    .TextMatrix(0, 5) = "„’«—Ì› ‘Õ‰"
    .TextMatrix(0, 6) = "„’«—Ì› »Ê·Ì’…"
    .TextMatrix(0, 7) = " ﬁÌ„… «·›« Ê—…"
    .TextMatrix(0, 8) = " «·≈Ã„«·Ï"
    .TextMatrix(0, 9) = "‘—ﬂ… «·‘Õ‰"
    .TextMatrix(0, 10) = "—ﬁ„ »Ê·Ì’…"
    .TextMatrix(0, 11) = "ﬂ«‘Ì—"
    .TextMatrix(0, 12) = "»‰ﬂ «·≈Ìœ«⁄"
    .TextMatrix(0, 13) = " «—ÌŒ  Õ’Ì·"
    .TextMatrix(0, 15) = "› Õ «·„” ‰œ"
    .TextMatrix(0, 16) = "—›÷ «·»Ê‰"
    .TextMatrix(0, 17) = "—ﬁ„ «·ÿ·»Ì…"
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 1300
    .ColWidth(2) = 1200
    .ColWidth(3) = 1800
    .ColWidth(4) = 1600
    .ColWidth(5) = 800
    .ColWidth(6) = 800
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    .ColWidth(9) = 1200
    .ColWidth(10) = 2000
    .ColWidth(11) = 2000
    .ColWidth(12) = 1000
    .ColWidth(13) = 1300
    .ColHidden(14) = True
    .ColWidth(15) = 700
    .ColWidth(16) = 2000
    .ColWidth(17) = 1000
     
    .ColComboList(15) = "..."
    .ColDataType(0) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(1) = flexDTDate
    .ColDataType(13) = flexDTDate
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ColDataType(8) = flexDTDouble
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 5, "#0", , , True, "«·«Ã„«·Ì "
    .Subtotal flexSTSum, -1, 6, "#0", , , True, "«·«Ã„«·Ì "
    .Subtotal flexSTSum, -1, 7, "#0", , , True, "«·«Ã„«·Ì "
    .Subtotal flexSTSum, -1, 8, "#0", , , True, "«·«Ã„«·Ì "
    .Subtotal flexSTCount, -1, 2, "#0", , , True, "«·«Ã„«·Ì "
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With grid1
    If .Col = 15 Then
        SalesFrm.sDoc_no = .TextMatrix(.Row, 14)
        SalesFrm.Show
    End If
End With
myload
End Sub
Private Sub grid1_EnterCell()
With grid1
If .Col = 15 Then
    .Editable = flexEDKbdMouse
Else
    .Editable = flexEDNone
End If
End With
End Sub
