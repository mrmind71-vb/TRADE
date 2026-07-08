VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form StockFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ã—œ „Œ«“‰"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14310
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
   ScaleHeight     =   8475
   ScaleWidth      =   14310
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Height          =   645
      Left            =   2070
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   945
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "«÷«›… «’‰«› «·„Œ“‰"
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
         TabIndex        =   39
         Top             =   135
         Width           =   2355
      End
   End
   Begin VB.TextBox xFilter2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3105
      RightToLeft     =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "»ÕÀ"
      Top             =   7515
      Width           =   1770
   End
   Begin VB.Frame Frame4 
      Height          =   645
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   -45
      Visible         =   0   'False
      Width           =   4515
      Begin VB.CommandButton cmdAdd 
         Caption         =   "«÷«›…"
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   225
         Width           =   960
      End
      Begin MSDataListLib.DataCombo xgroup 
         Height          =   315
         Left            =   1125
         TabIndex        =   33
         Top             =   225
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "„Ã„Ê⁄… :"
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
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.TextBox xfilter 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "»ÕÀ"
      Top             =   7515
      Width           =   1770
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   495
      ScaleWidth      =   14250
      TabIndex        =   19
      Top             =   7920
      Width           =   14310
      Begin VB.CommandButton cmdFix 
         BackColor       =   &H00E0E0E0&
         Caption         =   "«⁄«œ… ÷»ÿ «·Ã—œ"
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
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   90
         Width           =   1635
      End
      Begin VB.CommandButton cmdunPost 
         BackColor       =   &H00E0E0E0&
         Caption         =   "≈·€«¡  —ÕÌ·"
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
         Left            =   5490
         RightToLeft     =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   90
         Width           =   1500
      End
      Begin VB.CommandButton cmdPost 
         BackColor       =   &H00E0E0E0&
         Caption         =   " —ÕÌ·"
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
         Left            =   7020
         RightToLeft     =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   90
         Width           =   1950
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "|<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "«ŒÌ—"
         Top             =   90
         Width           =   465
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   510
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   " «·Ì"
         Top             =   90
         Width           =   465
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   975
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "”«»ﬁ"
         Top             =   90
         Width           =   465
      End
      Begin VB.CommandButton cmdfirst 
         Caption         =   ">|"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   90
         Width           =   465
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "«·≈Ã„«·Ì:"
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
         Left            =   13320
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   180
         Width           =   870
      End
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   11835
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   135
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   7290
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton Command2 
         Caption         =   "ÿ»«⁄… «·Ã—œ"
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
         Left            =   5445
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   135
         Width           =   1365
      End
      Begin VB.CommandButton cmdInform 
         Caption         =   "≈” ⁄·«„"
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
         Left            =   4095
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewinv 
         Caption         =   "„” ‰œ ÃœÌœ"
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
         Left            =   2745
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
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
         Left            =   45
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelInv 
         Caption         =   "Õ–› «·„” ‰œ"
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
         Left            =   1395
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   6030
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   585
      Width           =   8160
      Begin VB.TextBox xDesca 
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
         Left            =   3870
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   585
         Width           =   3165
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   585
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
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
         Height          =   330
         Left            =   5670
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1365
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
         Height          =   330
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   2040
      End
      Begin VB.Label Label4 
         Caption         =   "»Ì«‰ :"
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
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   630
         Width           =   705
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
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   315
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
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„Œ“‰ :"
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
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   675
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   4635
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   630
      Width           =   1380
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
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
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   540
         Width           =   1230
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ "
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
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   1230
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   2325
      Top             =   1725
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
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
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   2205
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   5130
      Top             =   2835
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Left            =   1125
      Top             =   3060
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   5820
      Left            =   90
      TabIndex        =   17
      Top             =   1620
      Width           =   14100
      _cx             =   24871
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   12648447
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
      Cols            =   9
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
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   330
      Left            =   7380
      TabIndex        =   29
      Top             =   7515
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "»ÕÀ ﬂÊœ :"
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
      Left            =   4995
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   7560
      Width           =   915
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "»ÕÀ Ê’› :"
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
      Left            =   1980
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   7560
      Width           =   915
   End
End
Attribute VB_Name = "StockFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CardTable As ADODB.Recordset, Search31 As New Search3, searchImp As New Search3
Dim clist1 As String, cList2 As String
Public bEdit As Boolean
Dim itemBalTable As ADODB.Recordset
Dim formMode, dDateLast As String
Dim con As New ADODB.Connection
Const LoadMode = 0, DefineMode = 1
Sub ItemsLookup()
Dim Generalarray(5)
Dim listarray(1, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select File1_10.item,File1_10.Desca,file1_50.desca,file1_10.cost From file1_10 left join file1_50 on file1_10.[GROUP] = file1_50.code"
Generalarray(2) = "Order by file1_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(FILE1_10.ITEM LIKE 'cFilter%' or  %%FILE1_10.DESCA%%) "

listarray(1, 0) = "«·„Ã„Ê⁄…"
listarray(1, 1) = "(%%FILE1_50.DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·’‰›"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "≈”„ «·’‰›"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "«· ﬂ·›…"
GrdArray(3, 1) = 0

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "«” ⁄·«„ «·«’‰«›"
Search3.Show 1
End Sub
Private Function myreplace() As Boolean
Dim aInsert(3, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = addDate(xDate.Text)

aInsert(2, 0) = "store"
aInsert(2, 1) = addstring(xStore.BoundText)

aInsert(3, 0) = "Desca"
aInsert(3, 1) = addstring(xdesca.Text)

con.BeginTrans
If xDoc_No.Enabled Then
    xDoc_No.Text = RetZero(Val(Newflag("FILE0_10h", "doc_no", con)))
    aInsert(0, 1) = addstring(xDoc_No.Text)
    con.Execute CreateInsert(aInsert, "FILE0_10h")
Else
    con.Execute CreateUpdate(aInsert, "FILE0_10h", " where doc_no = " & addstring(xDoc_No.Text))
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
Sub myproc()
If ActiveControl.Name = grid1.Name Then
    nFound = FoundOtheritem(Row, Col, Search3.grid1.TextMatrix(Search3.grid1.Row, 0))
    If nFound <> -1 Then
        MsgBox "«·’‰› „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & grid1.TextMatrix(nFound, 1)
        Exit Sub
    End If
    cItem = grid1.TextMatrix(grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 0) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 2) = "1"
    If Trim(cItem) <> Trim(grid1.TextMatrix(grid1.Row, 0)) Then calcRow grid1.Row, 0
    If grid1.Row = grid1.Rows - 1 Then
        grid1.AddItem ""
        grid1.Select grid1.Rows - 1, 0
    ElseIf grid1.Row = grid1.Rows - 2 Then
        grid1.Select grid1.Rows - 1, 0
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "doc_No = " & MyParn(Search31.grid1.TextMatrix(Search31.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    MyLoad
    Search31.Hide
ElseIf ActiveControl.Name = xdocImp.Name Then
    xdocImp.Text = searchImp.grid1.TextMatrix(searchImp.grid1.Row, 0)
    searchImp.Hide
End If
End Sub
Private Sub cmd_item_Click()
Load frmItem
frmItem.Show 1
End Sub
Private Sub cmdaddImp_Click()
If xStore.BoundText = "" Then
    MsgBox "≈Œ Ì«— «·„Œ“‰"
    Exit Sub
End If
Dim impTable As New ADODB.Recordset
cString = "select FILE7_60.*,file1_10.desca,file1_10.[unit] as unit2,file1_10.cost,file1_10.price,file1_10.price2,file1_10.package,file1_10.[group],file1_10.[section],file1_10.discount from file7_60 inner join file1_10 on file7_60.item = file1_10.item where doc_no = " & MyParn(xdocImp.Text) & " AND FILE7_60.STORE = " & MyParn(xStore.BoundText)
If XGROUP.BoundText <> "" Then cString = cString & " AND FILE1_10.[GROUP] = " & XGROUP.BoundText
impTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With grid1
.Rows = 1
Do Until impTable.EOF
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = impTable!Item
    .TextMatrix(.Rows - 1, 1) = impTable![FILE1_10.DESCA] & ""
    .TextMatrix(.Rows - 1, 2) = impTable!unit2 & ""
    .TextMatrix(.Rows - 1, 5) = Val(.TextMatrix(.Rows - 1, 4)) - Val(.TextMatrix(.Rows - 1, 3))
    .TextMatrix(.Rows - 1, 7) = Val(.TextMatrix(.Rows - 1, 5)) * Val(.TextMatrix(.Rows - 1, 6))
    impTable.MoveNext
Loop
.AddItem ""
End With
CalcTotals
impTable.Close
Set impTable = Nothing
End Sub

Private Sub CmdAdd_Click()
'If xGroup.BoundText = "" Then Exit Sub
Dim loctable As New ADODB.Recordset
cString = "Select FILE1_10.*  from (FILE1_10 INNER JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE) INNER JOIN FILE1_50G ON FILE1_50.[GROUP] = FILE1_50G.CODE"
If XGROUP.BoundText <> "" Then
    cString = cString & turn(cString) & "FILE1_50.[GROUP] = " & XGROUP.BoundText
End If
cString = cString & " ORDER BY FILE1_10.ITEM"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    nFound = grid1.FindRow(loctable!Item, , 0)
    If nFound = -1 Then
        grid1.TextMatrix(grid1.Rows - 1, 0) = loctable!Item
        grid1.TextMatrix(grid1.Rows - 1, 1) = loctable!DESCA & ""
        grid1.TextMatrix(grid1.Rows - 1, 5) = loctable!cost & ""
        grid1.AddItem ""
    End If
    loctable.MoveNext
Loop

If Not MYVALIDfix Then Exit Sub
prog1.Value = 0
prog1.Visible = True
For i = 1 To grid1.Rows - 2
     prog1.Value = Round(i / (grid1.Rows - 2), 2) * 100
    grid1.TextMatrix(i, 3) = RetItemBalance(grid1.TextMatrix(i, 0), xStore.BoundText, xDate.Text, con) & ""
Next
prog1.Visible = False
CalcTotals
End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From FILE0_10 where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From FILE0_10H where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    CardTable.Requery
    If (CardTable.EOF And CardTable.BOF) Then
        mydefine
    Else
        CardTable.Find "doc_no < " & MyParn(xDoc_No.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        MyLoad
    End If
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
Private Sub cmdFix_Click()
If Not MYVALIDfix Then Exit Sub
With grid1
'Time = Now
prog1.Value = 0
prog1.Visible = True
For i = 1 To grid1.Rows - 2
     prog1.Value = Round(i / (grid1.Rows - 2), 2) * 100
    .TextMatrix(i, 3) = RetItemBalance(.TextMatrix(i, 0), xStore.BoundText, xDate.Text, con) & ""
    .TextMatrix(i, 5) = Val(.TextMatrix(i, 4)) - Val(.TextMatrix(i, 3))
    .TextMatrix(i, 7) = Val(.TextMatrix(i, 5)) * Val(.TextMatrix(i, 6))
Next
prog1.Visible = False
If myreplace Then
'MsgBox DateDiff("s", nTime, Now)
    Inform " „ ÷»ÿ «·„” ‰œ »‰Ã«Õ"
End If
End With
End Sub
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO,DATE,CONVERT(VARCHAR(10),[DATE],111),FILE0_40.DESCA " & _
                  " FROM FILE0_10H INNER JOIN FILE0_40 ON FILE0_10H.Store = FILE0_40.CODE "

Generalarray(2) = "Order by [Date]"
Generalarray(3) = 4200
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ-«·„Œ“‰"
listarray(0, 1) = "@@Doc_No@@6 or  FILE0_40.DESCA Like '%cFilter%' OR " & _
                  "##[DATE]##"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·„Œ“‰"
GrdArray(3, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search31
Search31.Caption = "«” ⁄·«„"
Search31.Show 1
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
MyLoad
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
MyLoad
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    MyLoad
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    MyLoad
End If
End Sub
Private Sub cmdPost_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then MyLoad

On Error GoTo myerror
con.BeginTrans
cString = " UPDATE FILE0_10H SET FILE0_10H.closed = 1 WHERE FILE0_10H.DOC_NO = " & MyParn(xDoc_No.Text)
con.Execute cString
con.CommitTrans

CardTable.Requery
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
handleControls LoadMode
MsgBox " „  —ÕÌ· «·„” ‰œ »‰Ã«Õ"
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdNewInv_Click()
'CardTable.MoveLast
'xDoc_No.Text = RetZero(Val(CardTable!doc_no & ""))
mydefine
xDoc_No.SetFocus
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
CardTable.Requery
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
MyLoad
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
End Sub
Private Sub CmdUndo_Click()
If CardTable.EOF And CardTable.BOF Then
    grid1.Rows = 1
    grid1.AddItem ""
    Exit Sub
End If
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then
    CardTable.MoveLast
    MyLoad
Else
    MyLoad
End If
End Sub
Private Sub cmdunPost_Click()
'On Error GoTo Myerror
con.BeginTrans
cString = " UPDATE FILE0_10H SET FILE0_10H.closed = 0 WHERE FILE0_10H.DOC_NO = " & MyParn(xDoc_No.Text)
con.Execute cString
con.CommitTrans
CardTable.Requery

CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
handleControls LoadMode
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub

Private Sub Command3_Click()
Load PrintGrd
grid1.TextMatrix(grid1.Rows - 1, grid1.Cols - 1) = xtotal.Caption
grid1.TextMatrix(grid1.Rows - 1, 0) = "«·«Ã„«·Ì"
PrintGrd.doprint grid1, 1, True
PrintGrd.Show 1
grid1.TextMatrix(grid1.Rows - 1, grid1.Cols - 1) = ""
grid1.TextMatrix(grid1.Rows - 1, 0) = ""
End Sub

Private Sub Command1_Click()
If Trim(xStore.BoundText) = "" Then Exit Sub
nFirstRow = grid1.Rows - 2
FillStockItems Me.grid1, xStore.BoundText, con
For i = nFirstRow To grid1.Rows - 2
    calcRow nFirstRow, 0
Next
End Sub

Private Sub Command2_Click()
If MsgBox("ÿ»«⁄… »Ì«‰ »œÊ‰ √—œ’… ··Ã—œ", vbYesNo) = vbYes Then
    doprint "stock.rpt"
Else
    doprint "stock2.rpt"
End If
End Sub
Private Sub Form_Load()
'FixCost
openCon con
'clist1 = StrList("select * from file1_50 order by desca")
'cList2 = StrList("select * from file1_10SC order by desca")

Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT * FROM File0_10H  ORDER BY DOC_NO ", con, adOpenStatic, adLockReadOnly, adCmdText

data1.ConnectionString = strCon
data1.RecordSource = "FILE0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

data2.ConnectionString = strCon
data2.RecordSource = "FILE1_50g"
Set XGROUP.RowSource = data2
XGROUP.ListField = "Desca"
XGROUP.BoundColumn = "Code"

Set grid1.DataSource = data3
data3.ConnectionString = strCon

'data2.ConnectionString = strCon
'data2.RecordSource = "SELECT Distinct Doc_No, DESCA FROM FILE0_30"

'Set xList.RowSource = data2
'xList.ListField = "Desca"
'xList.BoundColumn = "Doc_No"

'With grid1
'    .Cols = 7
'    .Rows = 1
'    .Editable = flexEDKbd
'    .FormatString = "ﬂÊœ|" & "«·’‰‹‹‹‹‹‹›|" & "«·ÊÕœ…|" & "—’Ìœ ﬂÊ„»ÌÊ —|" & "—’Ìœ Ã—œ |" & "«·›—ﬁ |" & " ﬂ·›… |" & "«· ﬁÌ„"
'    .ColWidth(0) = 1200
'    .ColWidth(1) = 3200
'    .ColWidth(2) = 1300
'    .ColWidth(3) = 1300
'    .ColWidth(4) = 1300
'    .ColWidth(5) = 1300
'    .ColWidth(6) = 1300
'    .ColWidth(7) = 1300
'    .ColHidden(2) = True
'
'    .ColDataType(3) = flexDTDouble
'    .ColDataType(4) = flexDTDouble
'    .ColDataType(5) = flexDTDouble
'    .ColDataType(6) = flexDTDouble
'    .ColDataType(0) = flexDTString
'
'    .ColAlignment(0) = flexAlignRightCenter
'    .ColAlignment(1) = flexAlignRightCenter
'    .ColAlignment(2) = flexAlignRightCenter
'    .ColAlignment(3) = flexAlignRightCenter
'    .ColAlignment(4) = flexAlignRightCenter
'    .ColAlignment(5) = flexAlignRightCenter
'    .ColAlignment(6) = flexAlignRightCenter
'    .ColAlignment(7) = flexAlignRightCenter
'End With
'With Grid2
'    .Cols = 7
'    .Rows = 1
'    .Editable = flexEDKbd
'    .FormatString = "ﬂÊœ|" & "«·’‰‹‹‹‹‹‹›|" & "—’Ìœ Ã—œ |" & "«·›—ﬁ |" & " ﬂ·›… |" & "«· ﬁÌ„"
'End With
grid1.ColHidden(2) = True
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    MyLoad
Else
    mydefine
End If
End Sub
Sub dispProc()
formMode = dispMode
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Unload Search3
Unload Search31
If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set CardTable = Nothing
closeCon con
Err.Clear
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = 0 Then calcRow Row, Col
If Not validRow(Row) Then Exit Sub
If Row = grid1.Rows - 1 Then grid1.AddItem ""
CalcTotals
End Sub

Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 Then
    If Not validRow(OldRow) Then grid1.RemoveItem OldRow
End If
End Sub

Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
ItemsLookup
End Sub

Private Sub grid1_EnterCell()
If (grid1.Col = 4 Or grid1.Col = 0 Or grid1.Col = 6) Or grid1.Col >= 8 Then
  grid1.Editable = flexEDKbd
    SetKbLayout Lang_EN
Else
   grid1.Editable = flexEDNone
End If
End Sub

Private Sub Grid1_GotFocus()
If grid1.Row = 0 Then
    grid1.SetFocus
    grid1.Select 1, 2
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim cwhere As String
If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 And validRow(grid1.Row) Then
    grid1.AddItem "", grid1.Row
End If
If KeyCode = 112 And grid1.Col = 0 And grid1.Rows <> 0 Then
    cwhere = "FILE1_10.ISRAW = 1"
    If xStore.BoundText <> "" Then sGroup = GetDesca("Select [group] from file0_40 where code = " & MyParn(xStore.BoundText))
    If sGroup <> "" Then
        cwhere = cwhere & turn(cwhere, " and ") & "FILE1_50.[GROUP] = " & MyParn(sGroup)
    End If
    ItemsLookupAll Me, Search3, cwhere
End If
End Sub

Private Sub Grid1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If KeyAscii = 13 And grid1.Col = 0 Then
    If grid1.Row = grid1.Rows - 1 Then
        'grid1.AddItem ""
        grid1.Select grid1.Rows - 1, 1
    Else
        grid1.Select grid1.Row + 1, 1
    End If
End If
If KeyAscii = 13 Then
    Select Case Col
        Case 0
            grid1.Col = 4
            grid1.Row = Row
        Case 4
            grid1.Col = 0
            'grid1.Row = Row + 1
    End Select
End If

End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 0 Then ItemsLookup

If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And CmdSave.Enabled Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            con.Execute "delete from file0_10 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
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
Select Case grid1.Col
    Case 0
        If KeyCode = 27 Then Exit Sub
        If KeyCode = 112 Then ItemsLookup
End Select
End Sub
Private Function MYVALID() As Boolean
If foundOther Then Exit Function
If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xStore.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰"
    Exit Function
End If

If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
    Exit Function
End If


With grid1
prog1.Visible = True
prog1.Value = 0
For i = 1 To .Rows - 2
    prog1.Value = Round(i / (grid1.Rows - 2), 2) * 100
    If .TextMatrix(i, 0) = "" Then
        .Select i, 0, i, grid1.Cols - 1
        MsgBox "ﬂÊœ «·’‰› €Ì— „ÊÃÊœ"
        Exit Function
    Else
        If GetDesca("select item from file1_10 where file1_10.item = " & MyParn(.TextMatrix(i, 0))) = "" Then
            .Select i, 0, i, 2
            MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
            Exit Function
        End If
    End If
Next
prog1.Visible = False
End With
MYVALID = True
End Function
Private Sub MyLoad()
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xStore.BoundText = CardTable!store & ""
xdesca.Text = CardTable!DESCA & ""
cString = "SELECT FILE0_10.ITEM, FILE1_10.DESCA,FILE1_10.UNIT,FILE0_10.ComputerBal ,File0_10.RealBal,File0_10.differ,File0_10.COST,file0_10.Realbal * FILE0_10.COST as Total,ID  " & _
      " FROM FILE0_10 INNER JOIN FILE1_10 ON FILE0_10.ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(xDoc_No.Text) & " order by FILE0_10.ROW"
data3.RecordSource = cString
data3.Refresh
grid1.AddItem ""
handleControls LoadMode
CalcTotals
FixGrd
End Sub
Private Sub mydefine()
xDoc_No.Text = RetZero(Val(Newflag("FILE0_10h", "doc_no", con)))
xDate.Text = Format(Date, "dd-mm-yyyy")
xStore.BoundText = ""
xdesca.Text = ""
grid1.Rows = 1
grid1.AddItem ""
handleControls DefineMode
FixGrd
End Sub
Private Sub handleControls(nMode)
Dim bClosed As Boolean
'cmdAddList.Enabled = (grid1.Rows <= 2)
cmdNewinv.Enabled = nMode = LoadMode And bEdit
CmdFirst.Enabled = (nMode = LoadMode)
CmdLast.Enabled = (nMode = LoadMode)
CmdNext.Enabled = (nMode = LoadMode)
CmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
If Not (CardTable.EOF And CardTable.BOF) Then bClosed = CardTable!CLOSED

cmdPost.Enabled = (Not bClosed) And bEdit And nMode = LoadMode
cmdunPost.Enabled = bClosed And bEdit And nMode = LoadMode
cmdFix.Enabled = (Not bClosed) And bEdit And nMode = LoadMode
CmdSave.Enabled = (nMode = DefineMode Or Not bClosed) And bEdit
CmdDelInv.Enabled = (Not bClosed) And bEdit And nMode = LoadMod
End Sub

Private Sub grid1_Validate(Cancel As Boolean)
    If Not validRow(grid1.Row) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 Then grid1.RemoveItem grid1.Row
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
Dim cGroup As String
If Col = 0 And Trim(grid1.EditText) <> "" Then
    nFound = FoundOtheritem(Row, Col, Trim(grid1.EditText))
    
    If nFound <> -1 Then
        MsgBox "«·’‰› „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & grid1.TextMatrix(nFound, 1)
        Cancel = True
    End If
    
    If GetDesca("select item from file1_10 where israw = 1 and file1_10.item = " & MyParn(grid1.EditText)) = "" Then
        If Not CheckGroup(Row) Then
            MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
            Cancel = True
        End If
    End If
    
Else
    If grid1.TextMatrix(Row, 0) = "" Then
        MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
        Cancel = True
    End If
End If
End With
End Sub
Private Sub xdate_Change()
'cmdAddList.Enabled = xStore.BoundText <> "" And IsDate(xDate.Text)
End Sub

Private Sub xDoc_No_LostFocus()
If xDoc_No.Text = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text)
If (CardTable.EOF And CardTable.BOF) Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then MyLoad
End Sub
Private Function MYVALIDfix() As Boolean
If foundOther Then Exit Function
If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xStore.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰"
    Exit Function
End If

If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
    Exit Function
End If


With grid1
For i = 1 To .Rows - 2
    If .TextMatrix(i, 0) = "" Then
        .Select i, 0, i, grid1.Cols - 1
        MsgBox "ﬂÊœ «·’‰› €Ì— „ÊÃÊœ"
        Exit Function
    Else
        If GetDesca("SELECT ITEM FROM FILE1_10 WHERE ITEM = " & MyParn(.TextMatrix(i, 0))) = "" Then
            .Select i, 0, i, 2
            MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
            Exit Function
        End If
    End If
Next
End With
MYVALIDfix = True
End Function
Private Sub calcRow(nRow, nCol)
With grid1
.TextMatrix(nRow, 1) = ""
.TextMatrix(nRow, 2) = ""
.TextMatrix(nRow, 3) = ""
.TextMatrix(nRow, 5) = ""
If nCol = 0 Then .TextMatrix(nRow, 6) = ""
.TextMatrix(nRow, 7) = ""

If Trim(grid1.TextMatrix(nRow, 0)) = "" Then Exit Sub

Dim aret As Variant
aret = aGetDesca("Select Desca,Cost from file1_10 where file1_10.item = " & MyParn(grid1.TextMatrix(nRow, 0)))
If UBound(aret) > 0 Then
    .TextMatrix(nRow, 1) = aret(1) & ""
    If IsDate(xDate.Text) And Trim(xStore.BoundText) <> "" Then
        grid1.TextMatrix(nRow, 3) = RetItemBalance(grid1.TextMatrix(nRow, 0), xStore.BoundText, xDate.Text, con) & ""
    End If
    .TextMatrix(nRow, 5) = Val(.TextMatrix(nRow, 4)) - Val(.TextMatrix(nRow, 3)) & ""
    If nCol = 0 Then .TextMatrix(nRow, 6) = itemCost(grid1.TextMatrix(nRow, 0), xDate.Text) & ""
    .TextMatrix(nRow, 7) = Val(.TextMatrix(nRow, 4)) * Val(.TextMatrix(nRow, 6)) & ""
End If
End With
End Sub
Private Function RetItemBalance2(cItem, cStore, dDate) As Double
If cItem = "" Then Exit Function
cString = "Select sum(val([IN] & '') - VAL([OUT] & '')) as Balance From file1_11 where item = " & MyParn(cItem) & _
          " and Store = " & MyParn(cStore) & " and Date < " & DateSq(dDate)
RetItemBalance2 = Val(GetDesca(cString) & "")
End Function

Private Sub xdocimp_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then impcostLookup
End Sub

Private Sub xfilter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then FilterGrd xfilter.Text, 1
End Sub

Private Sub xFilter2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then FilterGrd xFilter2.Text, 0
End Sub

Private Sub xgroup_LostFocus()
If Not XGROUP.MatchedWithList Then XGROUP.BoundText = ""
End Sub

Private Sub xStore_Click(Area As Integer)
'cmdAddList.Enabled = xStore.BoundText <> "" And IsDate(xDate.Text)
End Sub
Private Function RetTotalCost()
For i = 1 To grid1.Rows - 1
    RetTotalCost = RetTotalCost + Val(grid1.TextMatrix(i, 0))
Next
End Function
Private Sub CalcTotals()
Dim nTotal As Double
With grid1
For i = 1 To grid1.Rows - 1
    .TextMatrix(i, 5) = Val(.TextMatrix(i, 4)) - Val(.TextMatrix(i, 3))
    nTotal = nTotal + (Val(grid1.TextMatrix(i, 4)) * Val(grid1.TextMatrix(i, 6)))
Next
End With
xtotal.Caption = Format(nTotal, "#0.00")
End Sub
Private Sub FixGrd()
With grid1
.FormatString = "ﬂÊœ|" & "«·’‰‹‹‹‹‹‹›|" & "«·ÊÕœ…|" & "—’Ìœ ﬂÊ„»ÌÊ —|" & "—’Ìœ Ã—œ |" & "«·›—ﬁ |" & " ﬂ·›… |" & "«· ﬁÌ„|"
.ColWidth(0) = 1500
.ColWidth(1) = 5000
.ColWidth(2) = 1000
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 1000
.ColWidth(6) = 1000
.ColWidth(7) = 1000
.ColWidth(8) = 1000

'.ColHidden(8) = True
'.ColComboList(12) = clist1
'.ColComboList(13) = cList2
.ColDataType(3) = flexDTDouble
.ColDataType(4) = flexDTDouble
.ColDataType(5) = flexDTDouble
.ColDataType(6) = flexDTDouble
.ColDataType(0) = flexDTString
For i = 0 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
'.ColHidden(2) = True

'.ColComboList(0) = "..."
.ColHidden(.Cols - 1) = True
End With
End Sub
Private Sub doprint(cReport)
Dim loctable As New ADODB.Recordset
If Not MYVALID Then Exit Sub
Dim TempTable As New ADODB.Recordset

loctable.Open "select file1_10.item,file1_10.[GROUP],file1_50.desca " & _
              " from file1_10 left join file1_50 on file1_10.[GROUP] = file1_50.code", con, adOpenStatic, adLockReadOnly, adCmdText
contemp.Execute "DELETE * FROM TEMP"
TempTable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

For i = 1 To grid1.Rows - 2
    TempTable.AddNew
    TempTable!date3 = DateFix(xDate.Text)
    TempTable!STR21 = turn(xdesca.Text, "", "Ã—œ ") & Trim(xdesca.Text) & turn(xdesca.Text, " ") & "·„Œ“‰ " & xStore.Text
    TempTable!str1 = TurnValue(grid1.TextMatrix(i, 0))
    TempTable!str2 = TurnValue(grid1.TextMatrix(i, 1))
    TempTable!VAL1 = Val(grid1.TextMatrix(i, 3))
    TempTable!VAL2 = Val(grid1.TextMatrix(i, 4))
    TempTable!VAL3 = Val(grid1.TextMatrix(i, 5))
    TempTable!VAL4 = Val(grid1.TextMatrix(i, 6))
    TempTable!VAL5 = Val(grid1.TextMatrix(i, 7))
    TempTable!VAL11 = i
    TempTable.Update
Next
If TempTable.EOF And TempTable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\" & cReport
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
TempTable.Close
Set TempTable = Nothing
End Sub

Private Function foundOther() As Boolean
For i = 1 To grid1.Rows - 2
    nRow = FoundOtherRow(i, 0)
    If nRow <> -1 Then
        MsgBox "«·’‰› " & grid1.TextMatrix(nRow, 1) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ " & i & " Ê " & nRow
        grid1.Select nRow, 0, nRow, grid1.Cols - 1
        If Not grid1.RowIsVisible(nRow) Then grid1.ShowCell nRow, 0
        foundOther = True
        Exit Function
    End If
Next
End Function
Private Function FoundOtherRow(nRow, nCol) As Integer
FoundOtherRow = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = Trim(grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = i
            Exit Function
        End If
    End If
Next
End Function
Private Function FoundOtheritem(nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = nValue Then
            FoundOtheritem = i
            Exit Function
        End If
    End If
Next
End Function
Private Sub myreplaceGrd()
Dim nCost As Double
Dim aInsert(6, 1)
With grid1
    prog1.Value = 0
    prog1.Visible = True
    For i = 1 To .Rows - 2
        prog1.Value = Round(i / (grid1.Rows - 2), 2) * 100
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.Text)
        
        aInsert(1, 0) = "item"
        aInsert(1, 1) = addstring(grid1.TextMatrix(i, 0))
        
        aInsert(2, 0) = "ComputerBal"
        aInsert(2, 1) = Val(.TextMatrix(i, 3))

        aInsert(3, 0) = "RealBal"
        aInsert(3, 1) = Val(.TextMatrix(i, 4))

        aInsert(4, 0) = "Differ"
        aInsert(4, 1) = Val(.TextMatrix(i, 5))

        aInsert(5, 0) = "COST"
        aInsert(5, 1) = Val(.TextMatrix(i, 6))

        aInsert(6, 0) = "row"
        aInsert(6, 1) = i
        
'        If grid1.TextMatrix(i, .Cols - 1) = "" Then
'            nCost = Val(grid1.TextMatrix(i, 6))
'            aInsert(6, 0) = "cost"
'            aInsert(6, 1) = nCost
'        End If
        
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, "FILE0_10")
        Else
            con.Execute CreateUpdate(aInsert, "FILE0_10", " where ID = " & grid1.TextMatrix(i, .Cols - 1), Array(0, 6))
        End If
    Next
 prog1.Visible = False
End With
End Sub
Private Function validRow(nRow) As Boolean
If Trim(grid1.TextMatrix(nRow, 0)) = "" Then Exit Function
validRow = True
End Function
Private Function CheckGroup(nRow As Long) As Boolean
With grid1
cColor = mySplit(.EditText, 3)
cExit = mySplit(.EditText, 4)
If cExit <> "" Then Exit Function
If cColor = "" Then Exit Function
cColorName = GetDesca("select desca from color where code = " & MyParn(cColor))
If cColorName <> "" Then
    cGroup = mySplit(.EditText, 1) & "-" & mySplit(.EditText, 2)
    cGroupname = GetDesca("SELECT DESCA FROM FILE1_50 WHERE CODE = " & MyParn(cGroup))
    If cGroupname = "" Then Exit Function
    If MsgBox("«÷«›… ’‰› " & cGroupname & " " & cColorName, vbOKCancel + vbDefaultButton2) = vbOK Then
        cString = "insert into file1_10(ITEM,DESCA,SUBITEM,DESCASUB,[GROUP],ISRAW,[COLOR])" & _
                  " VALUES(" & _
                  addstring(cGroup & "-" & cColor) & "," & _
                  addstring(cGroupname & " " & cColorName) & "," & _
                  addstring(cColor) & "," & _
                  addstring(cColorName) & "," & _
                  addstring(cGroup) & "," & _
                  "1" & "," & _
                  addstring(cColor) & _
                  ")"
        On Error GoTo myerror
        con.BeginTrans
        con.Execute cString
        con.CommitTrans
        .TextMatrix(nRow, 1) = cGroupname & " " & cColorName
        CheckGroup = True
    End If
End If
End With
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Function

Private Sub xgroup_GotFocus()
XGROUP.BackColor = &HC0FFFF
End Sub
Private Sub xDescA_GotFocus()
xdesca.SelStart = 0
xdesca.SelLength = Len(xdesca.Text)
End Sub
Private Sub xStore_GotFocus()
xStore.BackColor = &HC0FFFF
End Sub
Private Sub xDoc_No_GotFocus()
xDoc_No.SelStart = 0
xDoc_No.SelLength = Len(xDoc_No.Text)
End Sub
Private Sub xDate_GotFocus()
xDate.SelStart = 0
xDate.SelLength = Len(xDate.Text)
End Sub
Private Sub FixCost()
If RetSetting("alternew2", App.Path & "\alter.txt") = "yes" Then Exit Sub
Dim FS1 As ADODB.Command
Set FS1 = New ADODB.Command

openCon con

Dim loctable As New ADODB.Recordset

loctable.Open "SELECT FILE0_10.ITEM,FILE1_10.[GROUP],FILE1_10.SUBITEM FROM FILE0_10 INNER JOIN FILE1_10 ON FILE0_10.ITEM = FILE1_10.ITEM ", con, adOpenStatic, adLockReadOnly
Do Until loctable.EOF
    con.Execute "UPDATE FILE0_10 SET FILE0_10.ITEM = " & addstring(loctable!Group & "-" & loctable!SUBITEM) & _
                " WHERE FILE0_10.ITEM = " & MyParn(loctable!Item)
    loctable.MoveNext
Loop

'On Error GoTo myerror
cString = "UPDATE FILE1_10 SET FILE1_10.ITEM = FILE1_10.[GROUP] + '-' +  FILE1_10.SUBITEM"
cString = cString & turn(cString) & " FILE1_10.ISRAW = 1"
cString = cString & turn(cString) & " (NOT (FILE1_10.SUBITEM IS NULL))"

FS1.CommandType = adCmdText
Set FS1.ActiveConnection = con
FS1.CommandText = cString
FS1.Execute
Set FS1 = Nothing


Set FS1 = New ADODB.Command
cString = "ALTER TABLE [dbo].[FILE0_10] ADD [COST2] [decimal](18, 3) NULL"
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = con
FS1.CommandText = cString
FS1.Execute
Set FS1 = Nothing

Set FS1 = New ADODB.Command
cString = "UPDATE FILE0_10 SET COST2 = COST"
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = con
FS1.CommandText = cString
FS1.Execute

closeCon con
addSetting "alternew2", "yes", App.Path & "\alter.txt"
MsgBox " „ ÷»ÿ «·„·› »‰Ã«Ã"
GoTo ClearVar:
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
ClearVar:
closeCon con
Set fs = Nothing
End Sub
Private Sub FilterGrd(cText, nCol)
Dim bHide As Boolean
For i = 1 To grid1.Rows - 2
     grid1.RowHidden(i) = False
Next
If Trim(cText) = "" Then Exit Sub

Dim aret As Variant
aret = Split(Trim(cText))
For i = 1 To grid1.Rows - 2
    For n = 0 To UBound(aret)
        bHide = InStr(1, Trim(grid1.TextMatrix(i, nCol)), aret(n)) = 0
        If bHide Then Exit For
    Next
    grid1.RowHidden(i) = bHide
Next
End Sub

