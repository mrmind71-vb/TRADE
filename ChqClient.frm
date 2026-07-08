VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form chqClientfrm 
   Caption         =   "√Ê—«ﬁ ﬁ»÷"
   ClientHeight    =   7905
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   9450
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
   ScaleHeight     =   7905
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox xDateBank 
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
      Left            =   5130
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   3465
      Width           =   2715
   End
   Begin VB.TextBox xdesca 
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
      Height          =   1050
      Left            =   90
      MaxLength       =   50
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2745
      Width           =   3300
   End
   Begin VB.CheckBox xOld 
      Alignment       =   1  'Right Justify
      Caption         =   "‘Ìﬂ ”«»ﬁ"
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
      Left            =   3315
      RightToLeft     =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   675
      Width           =   2115
   End
   Begin VB.TextBox xName2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2565
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1665
      Width           =   4020
   End
   Begin VB.TextBox xCode2 
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
      Left            =   6615
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1665
      Width           =   1230
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   9450
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   0
      Width           =   9450
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqClient.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   5940
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqClient.frx":014A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Õ›Ÿ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdDel 
         Height          =   420
         Left            =   4800
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqClient.frx":058C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   420
         Left            =   3660
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqClient.frx":2E26
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdAdd 
         Height          =   420
         Left            =   7080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqClient.frx":539F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   8220
         Picture         =   "ChqClient.frx":794B
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   45
         Width           =   1140
      End
      Begin VB.CommandButton cmdSum 
         Caption         =   "‘Ìﬂ«  „Ã„⁄…"
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
         Picture         =   "ChqClient.frx":A11E
         RightToLeft     =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   45
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1950
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   9450
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   7455
      Width           =   9450
      Begin VB.OptionButton optclose 
         BackColor       =   &H80000010&
         Caption         =   "«·ﬂ‹‹‹·"
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
         Index           =   3
         Left            =   2415
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton optclose 
         BackColor       =   &H80000010&
         Caption         =   "€Ì— „Õ’·…"
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
         Index           =   0
         Left            =   6300
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   135
         Width           =   1215
      End
      Begin VB.OptionButton optclose 
         BackColor       =   &H80000010&
         Caption         =   "„Õ’·…"
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
         Index           =   2
         Left            =   5085
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   135
         Width           =   990
      End
      Begin VB.OptionButton optclose 
         BackColor       =   &H80000010&
         Caption         =   "„—›Ê÷…"
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
         Index           =   1
         Left            =   3690
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   135
         Width           =   1215
      End
   End
   Begin VB.TextBox xCode1 
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
      Left            =   6615
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1305
      Width           =   1230
   End
   Begin VB.TextBox xNAME1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   2565
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1305
      Width           =   4020
   End
   Begin VB.TextBox xDATE_1 
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
      Left            =   5130
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2025
      Width           =   2715
   End
   Begin VB.TextBox xBANK_REC 
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
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2385
      Width           =   3300
   End
   Begin VB.TextBox xNAME4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   90
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2025
      Width           =   3300
   End
   Begin VB.TextBox xValue 
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
      Left            =   5130
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2745
      Width           =   2715
   End
   Begin VB.TextBox XSER_NO 
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
      Left            =   5490
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   585
      Width           =   2355
   End
   Begin VB.TextBox XCHK_ID 
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
      Left            =   5490
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   945
      Width           =   2355
   End
   Begin VB.TextBox xDATE_R 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   5130
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2385
      Width           =   2715
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -2115
      Top             =   585
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
   Begin MSDataListLib.DataCombo XID_BANK 
      Height          =   315
      Left            =   5130
      TabIndex        =   7
      Top             =   3105
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   "DataCombo1"
      RightToLeft     =   -1  'True
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   8700
      Top             =   3975
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   -2115
      Top             =   990
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
   Begin VB.Frame Frame1 
      Height          =   1410
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   4275
      Width           =   7845
      Begin VB.TextBox xDATE_3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3300
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   600
         Width           =   1890
      End
      Begin VB.TextBox xMEMO 
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
         Left            =   225
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   945
         Width           =   4965
      End
      Begin VB.OptionButton xClosed 
         Alignment       =   1  'Right Justify
         Caption         =   "‘Ìﬂ €Ì— „Õ’·"
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
         Height          =   345
         Index           =   0
         Left            =   5250
         RightToLeft     =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   150
         Value           =   -1  'True
         Width           =   1740
      End
      Begin VB.OptionButton xClosed 
         Alignment       =   1  'Right Justify
         Caption         =   " ŸÂÌ— &  Õ’Ì· "
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
         Height          =   345
         Index           =   2
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   135
         Width           =   1740
      End
      Begin VB.OptionButton xClosed 
         Alignment       =   1  'Right Justify
         Caption         =   "—›÷ / —œ «·‘Ìﬂ"
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
         Height          =   345
         Index           =   1
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   150
         Width           =   1740
      End
      Begin MSDataListLib.DataCombo XBOX 
         Height          =   315
         Left            =   225
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Œ“«‰… :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ«  :"
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
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1050
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «·”œ«œ/  ŸÂÌ— / —›÷"
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
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   675
         Width           =   2145
      End
   End
   Begin VB.Frame frmTrans 
      Height          =   1095
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   5715
      Width           =   7845
      Begin VB.TextBox xTransName2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   675
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   585
         Width           =   2340
      End
      Begin VB.TextBox xTransCode2 
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
         Left            =   4365
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   585
         Width           =   1200
      End
      Begin VB.TextBox xTransCode1 
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
         Left            =   4365
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   1200
      End
      Begin VB.TextBox xTransName1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   675
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   225
         Width           =   2340
      End
      Begin VB.Label xLab3 
         AutoSize        =   -1  'True
         Caption         =   " ŸÂÌ— «·‘Ìﬂ ≈·Ï «·⁄„Ì·"
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
         Left            =   5745
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   630
         Width           =   1905
      End
      Begin VB.Label xLab4 
         AutoSize        =   -1  'True
         Caption         =   "≈”„ «·⁄„Ì·"
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
         Left            =   3225
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   690
         Width           =   915
      End
      Begin VB.Label xLab2 
         AutoSize        =   -1  'True
         Caption         =   "≈”„ «·„Ê—œ"
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
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   225
         Width           =   885
      End
      Begin VB.Label xLab1 
         AutoSize        =   -1  'True
         Caption         =   " ŸÂÌ— «·‘Ìﬂ ≈·Ï «·„Ê—œ"
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
         Left            =   5775
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   225
         Width           =   1875
      End
   End
   Begin VB.Frame Frame6 
      Height          =   555
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   6840
      Width           =   3300
      Begin VB.CommandButton cmdFirst 
         Height          =   375
         Left            =   45
         Picture         =   "ChqClient.frx":A650
         Style           =   1  'Graphical
         TabIndex        =   64
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdLast 
         Height          =   375
         Left            =   810
         Picture         =   "ChqClient.frx":CD2F
         Style           =   1  'Graphical
         TabIndex        =   63
         TabStop         =   0   'False
         ToolTipText     =   "√ŒÌ—"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   375
         Left            =   1755
         Picture         =   "ChqClient.frx":F409
         Style           =   1  'Graphical
         TabIndex        =   62
         TabStop         =   0   'False
         ToolTipText     =   "«·”«»ﬁ"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Height          =   375
         Left            =   2520
         Picture         =   "ChqClient.frx":119DC
         Style           =   1  'Graphical
         TabIndex        =   61
         TabStop         =   0   'False
         ToolTipText     =   "«· «·Ì"
         Top             =   135
         Width           =   735
      End
   End
   Begin MSDataListLib.DataCombo xacc 
      Height          =   315
      Left            =   3555
      TabIndex        =   65
      Top             =   3870
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSAdodcLib.Adodc DATA7 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Õ”«» : "
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
      Left            =   7965
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   3915
      Width           =   660
   End
   Begin VB.Label lblDateBank 
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ «·«Ìœ«⁄ :"
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   3420
      Width           =   1035
   End
   Begin VB.Label LabelCode2 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   1665
      Width           =   885
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ﬂÊœ «·„Ê—œ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   75
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "»‰ﬂ «·«Ìœ«⁄ :"
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   3105
      Width           =   915
   End
   Begin VB.Label label20 
      AutoSize        =   -1  'True
      Caption         =   "«·»Ì«‰ :"
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
      Left            =   3465
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   2745
      Width           =   525
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "„ŸÂ— „‰ :"
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
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   2070
      Width           =   795
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "«·»‰ﬂ «·„”ÕÊ» ⁄·ÌÂ :"
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
      Left            =   3465
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2430
      Width           =   1680
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ :"
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   2025
      Width           =   1380
   End
   Begin VB.Label LabelCode 
      AutoSize        =   -1  'True
      Caption         =   "ﬂÊœ «·⁄„Ì· :"
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1350
      Width           =   915
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "«·ﬁÌ„… :"
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   2745
      Width           =   570
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "—ﬁ„ «·‘Ìﬂ :"
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   990
      Width           =   945
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "„”·”· ‘Ìﬂ :"
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   630
      Width           =   1215
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ  Õ—Ì— :"
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   2340
      Width           =   945
   End
End
Attribute VB_Name = "chqClientfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim bSumMode As Boolean
Public sSer_no As String
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
CmdAdd.Enabled = (nMode = LoadMode And optclose(0).Value) And bedit
CmdSave.Enabled = (nMode = LoadMode Or optclose(0).Value) And bedit
cmdSum.Enabled = (nMode = LoadMode Or optclose(0).Value) And bedit
CmdUndo.Enabled = (nMode = LoadMode Or optclose(0).Value)
CmdDel.Enabled = (nMode = LoadMode) And bedit
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
xTransCode1.Enabled = (optclose(2).Value)
xTransCode2.Enabled = (optclose(2).Value)
XSER_NO.Enabled = Not (nMode = LoadMode)
End Sub
Sub CardLookup()
Dim Generalarray(5)
Dim listarray(2, 7)
Dim GrdArray(6, 1)

If optclose(0).Value Then cWhere = "CLOSED = '0'"
If optclose(1).Value Then cWhere = "CLOSED = '1'"
If optclose(2).Value Then cWhere = "CLOSED = '2'"

Set Generalarray(0) = Me
Generalarray(1) = " SELECT FILE5_20.SER_NO,FILE5_20.DESCA1,CONVERT(VARCHAR(10),FILE5_20.DATE_1 ,111), CONVERT(VARCHAR(10),FILE5_20.DATE_r ,111), [VALUE] , CHK_ID, FILE5_20.TRANSNAME1 From FILE5_20 " & turn(cWhere, " where ") & cWhere
Generalarray(2) = ""
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "„”·”· «Ê „” ›Ìœ -  «—ÌŒ  Õ—Ì—-—ﬁ„ «·‘Ìﬂ"
listarray(0, 1) = "(%%desca1%% or chk_id Like '%cFilter%' Or ##Date_R##)"

listarray(1, 0) = " «—ÌŒ «” Õﬁ«ﬁ"
listarray(1, 1) = " ##Date_1##"


listarray(2, 0) = "«·ﬁÌ„…"
listarray(2, 1) = "**[value]**"

GrdArray(0, 0) = "„”·”·"
GrdArray(0, 1) = 800

GrdArray(1, 0) = "‘Ìﬂ „‰"
GrdArray(1, 1) = 2000

GrdArray(2, 0) = "«” Õﬁ«ﬁ"
GrdArray(2, 1) = 1000

GrdArray(3, 0) = " Õ—Ì—"
GrdArray(3, 1) = 1000

GrdArray(4, 0) = "ﬁÌ„…"
GrdArray(4, 1) = 1000

GrdArray(5, 0) = "—ﬁ„ «·‘Ìﬂ"
GrdArray(5, 1) = 1200

GrdArray(6, 0) = " ŸÂÌ—"
GrdArray(6, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "«” ⁄·«„ «Ê—«ﬁ ﬁ»÷"
Search3.Show 1
End Sub
Sub CLIENTLOOKUP(Optional nFlag As Integer = 1)
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
If nFlag = 1 Then
    Generalarray(1) = "SELECT Code , Desca  From file3_10"
Else
    Generalarray(1) = "SELECT Code , Desca  From file4_10"
End If
Generalarray(2) = " Order by Code"
Generalarray(3) = 4000
Generalarray(5) = False

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·«”„"
GrdArray(1, 1) = 3000

listarray(0, 0) = "«·«”„"
listarray(0, 1) = "(%%desca%%)"

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "«” ⁄·«„" & " " & IIf(nFlag = 1, "«·⁄„·«¡", "«·„Ê—œÌ‰")
Search3.Show 1
End Sub
Sub mydefine()


XSER_NO.text = RetZero(Val(Newflag("FILE5_20", "ser_no", con)), 6)
XCHK_ID.text = ""
xOld.Value = 0
xCode1.text = ""
xCode2.text = ""
XID_BANK.BoundText = ""
xNAME1.text = ""
xName2.text = ""
xNAME4.text = ""
xBANK_REC.text = ""
xDATE_1.text = ""
xDATE_3.text = ""
xDATE_R.text = ""
XBOX.BoundText = ""
xacc.BoundText = ""
xValue.text = ""
xMEMO.text = ""
xdesca.text = ""
xClosed(0) = True
xClosed(1) = False
xClosed(2) = False
xTransCode1.text = ""
xTransCode2.text = ""
xTransName1.text = ""
xTransName2.text = ""
xDateBank.text = ""
Handlecontrols DefineMode
End Sub
Sub myDefine2()


XSER_NO.text = RetZero(Val(Newflag("FILE5_20", "ser_no", con)), 6)
XCHK_ID.text = Val(XCHK_ID.text) + 1
xOld.Value = 0
xName2.text = ""
xNAME4.text = ""
xDATE_1.text = ""
xValue.text = ""
xMEMO.text = ""
xdesca.text = ""
xClosed(0) = True
xClosed(1) = False
xClosed(2) = False
xTransCode1.text = ""
xTransCode2.text = ""
xTransName1.text = ""
xTransName2.text = ""
xDateBank.text = ""
Handlecontrols DefineMode
End Sub
Sub myProc()
If TypeOf ActiveControl Is TextBox Then
    ActiveControl.text = Search3.Grid1.TextMatrix(Search3.Grid1.Row, 0)
    Unload Search3
Else
    CardTable.Find "SER_NO = " & MyParn(Search3.Grid1.TextMatrix(Search3.Grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload
    Unload Search3
End If
End Sub
Sub myload()
XSER_NO.text = CardTable!Ser_no
XCHK_ID.text = CardTable!CHK_ID & ""
xCode1.text = CardTable!Code1 & ""
xCode2.text = CardTable!CODE2 & ""
xNAME1.text = CardTable!desca1 & ""
xName2.text = CardTable!Desca2 & ""
xNAME4.text = CardTable!NAME4 & ""
xdesca.text = CardTable!DESCA & ""
xClosed(0).Value = IIf(CardTable!CLOSED = "0", True, False)
xClosed(1).Value = IIf(CardTable!CLOSED = "1", True, False)
xClosed(2).Value = IIf(CardTable!CLOSED = "2", True, False)

xOld.Value = IIf(CardTable!OLD, 1, 0)
xBANK_REC.text = CardTable!Bank_rec & ""
xValue.text = Format(CardTable!Value, "Fixed")
xDATE_1.text = Format(CardTable!date_1, "dd-mm-yyyy")
xDATE_3.text = Format(CardTable!date_3, "dd-mm-yyyy")
xDATE_R.text = Format(CardTable!date_R, "dd-mm-yyyy")
xDateBank.text = Format(CardTable!DateBank, "dd-mm-yyyy")
xMEMO.text = CardTable!Memo & ""
XID_BANK.BoundText = CardTable!ID_BANK & ""
XBOX.BoundText = CardTable!BOX & ""
xacc.BoundText = CardTable!ACC & ""
xTransCode1.text = CardTable!TransCode1 & ""
xTransCode2.text = CardTable!TRANSCODE2 & ""

xTransName1.text = CardTable!TRANSNAME1 & ""
xTransName2.text = CardTable!transname2 & ""

Handlecontrols LoadMode
End Sub
Function MYVALID() As Boolean
If XSER_NO.text = "" Then
    MsgBox "ÌÃ»  ”ÃÌ· „””·”· ··‘Ìﬂ"
    Exit Function
End If
If Not IsDate(xDATE_R.text) Then
    MsgBox "ÌÃ»  ”ÃÌ·  «—ÌŒ «· Õ—Ì—"
    Exit Function
End If
If Me.xTransCode1.text <> "" And Me.xTransCode2.text <> "" Then
    MsgBox " „  ”ÃÌ· ⁄„Ì· Ê „Ê—œ ·‰›” «·‘Ìﬂ"
    Exit Function
End If

If xCode1.text <> "" Then
    If GetDesca("select code from file3_10 where code = " & MyParn(xCode1.text), con) = "" Then Exit Function
End If

If xCode2.text <> "" Then
    If GetDesca("select code from file4_10 where code = " & MyParn(xCode2.text), con) = "" Then Exit Function
End If

If (xTransCode1.text <> "" Or xTransCode2.text <> "" Or Not xClosed(0).Value) And Not IsDate(xDATE_3.text) Then
    MsgBox " ”ÃÌ·  «—ÌŒ «· ŸÂÌ— °  ÕœÌœ «‰ «·‘Ìﬂ  „  ŸÂÌ—…"
    Exit Function
End If

If XSER_NO.Enabled And Trim(xCode1.text) <> "" Then
    cString = GetDesca("select ser_no from FILE5_20 where code1 = " & MyParn(xCode1.text) & " and Chk_Id = " & MyParn(XCHK_ID.text), con) & ""
    If Trim(cString) <> "" Then
        MsgBox "‘Ìﬂ »‰›” «·—ﬁ„ ·‰›” " & "«·⁄„Ì· "
        Exit Function
    End If
End If
MYVALID = True
End Function
Private Sub CmdAdd_Click()
mydefine
On Error Resume Next
XCHK_ID.SetFocus
Err.Clear
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
con.BeginTrans
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    con.Execute "delete  From FILE5_20 where Ser_No = " & MyParn(XSER_NO.text)
End If
con.CommitTrans

CardTable.Requery
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.Find "SER_NO < " & MyParn(XSER_NO.text), , adSearchBackward, adBookmarkLast
    If CardTable.EOF Then CardTable.MoveFirst
    myload
Else
    If optclose(0).Value Then CmdAdd_Click Else mydefine
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
Private Sub cmdSave_Click()
'msgBoxStr = IIf(addmove, "«÷«›… ”Ã· : Â· «‰  „Ê«›ﬁ ø", "Õ›Ÿ «· €ÌÌ—«  ! Â· √‰  „Ê«›ﬁ ø")
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
CardTable.Requery
If XSER_NO.Enabled Then
    If bSumMode Then cmdSum_Click Else CmdAdd_Click
Else
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "SER_NO = " & MyParn(XSER_NO.text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then
            CardTable.Find "SER_NO < " & MyParn(XSER_NO.text), , adSearchBackward, adBookmarkLast
            If CardTable.BOF Then CardTable.MoveFirst
        End If
        myload
    Else
        mydefine
    End If
End If
End Sub
Private Sub CmdUndo_Click()
bSumMode = False
If CardTable.EOF And CardTable.BOF Then
    mydefine
Else
    If XSER_NO.Enabled Then
        CardTable.MoveLast
    Else
        CardTable.Find "ser_no = " & MyParn(XSER_NO.text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub cmdSum_Click()
myDefine2
On Error Resume Next
xDATE_1.SetFocus
Err.Clear
bSumMode = True
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Form_Load()
bedit = True
lblDateBank.Visible = True
xDateBank.Visible = True
    
openCon con
data1.ConnectionString = strCon
data1.RecordSource = "FILE0_50"

Set XBOX.RowSource = data1
XBOX.ListField = "Desca"
XBOX.BoundColumn = "Code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "FILE5_10"

Set XID_BANK.RowSource = DATA2
XID_BANK.ListField = "Desca"
XID_BANK.BoundColumn = "code"

DATA7.ConnectionString = strCon
DATA7.RecordSource = "SELECT * FROM acc0_10 ORDER BY DESCA"
Set xacc.RowSource = DATA7
xacc.ListField = "DESCA"
xacc.BoundColumn = "CODE"

optclose(0).Value = True
'If Not (CardTable.EOF And CardTable.BOF) Then
'    CardTable.MoveLast
'    MyLoad
'Else
'    CmdAdd_Click
'End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub
Private Sub optclose_Click(Index As Integer)
cString = "Select FILE5_20.*,FILE3_10.DESCA AS DESCA1,FILE4_10.DESCA AS DESCA2" & _
          " From (FILE5_20 LEFT JOIN FILE3_10 ON FILE5_20.CODE1 = FILE3_10.CODE) LEFT JOIN file4_10 ON FILE5_20.CODE2 = FILE4_10.CODE"
If Index <> 3 Then
    cString = cString & turn(cString) & " closed = " & MyParn(Index)
End If
cString = cString & " Order by Ser_No"

Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenKeyset, adLockOptimistic, adCmdText

If sSer_no <> "" And Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.Find "ser_no = " & MyParn(sSer_no), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload
    Exit Sub
End If

If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    If optclose(0).Value Then
        CmdAdd_Click
    Else
        mydefine
    End If
End If
End Sub
Private Sub XBOX_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then XBOX.BoundText = ""
End Sub

Private Sub xBOX2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xBox2.BoundText = ""
End Sub
Private Sub XCHK_ID_Validate(Cancel As Boolean)
If GetDesca("SELECT SER_NO FROM FILE5_20 WHERE SER_NO <> " & MyParn(XSER_NO.text) & " AND  CHK_ID = " & MyParn(XCHK_ID.text), con) <> "" Then
    MsgBox "—ﬁ„ «·‘Ìﬂ „ﬂ—— „‰ ﬁ»·"
End If
End Sub
Private Sub xClosed_Click(Index As Integer)
xTransCode1.Enabled = (Index = 2)
xTransCode2.Enabled = (Index = 2)
If Index <> 2 Then
    xTransCode1.text = ""
    xTransCode2.text = ""
    xTransName1.text = ""
    xTransName2.text = ""
End If
If Index = 0 Then
    xDATE_3.text = ""
End If
End Sub
Private Sub xCode1_Change()
If xCode1.text <> "" Then
    xCode2.text = ""
    xCode2.Enabled = False
    xName2.text = ""
Else
    xCode2.Enabled = True
End If
End Sub

Private Sub xCode1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CLIENTLOOKUP
End Sub
Private Sub xCode1_LostFocus()
xNAME1.text = ""
If Trim(xCode1.text) = "" Then Exit Sub
xCode1.text = RetZero(xCode1.text, 3)
xNAME1.text = GetDesca("Select Desca from FILE3_10 where code = " & MyParn(xCode1.text), con)
End Sub
Private Sub xCode2_Change()
If xCode2.text <> "" Then
    xCode1.text = ""
    xCode1.Enabled = False
    xNAME1.text = ""
Else
    xCode1.Enabled = True
End If
End Sub

Private Sub xCode2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CLIENTLOOKUP 2
End Sub
Private Sub xCODE2_LostFocus()
xName2.text = ""
If Trim(xCode2.text) = "" Then Exit Sub
xCode2.text = RetZero(xCode2.text, 3)
xName2.text = GetDesca("Select Desca from FILE4_10 where code = " & MyParn(xCode2.text), con)
End Sub
Private Sub XID_BANK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then XID_BANK.BoundText = ""

End Sub
Private Sub xSer_no_LostFocus()
If Trim(XSER_NO.text) = "" Then Exit Sub
XSER_NO.text = RetZero(XSER_NO.text, 6)
CardTable.Find "SER_NO = " & MyParn(XSER_NO.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Function myRecordCount() As Integer
'If RecordCountTable.RecordCount = 0 Then Exit Function
'RecordCountTable.MoveLast
'myRecordCount = RecordCountTable.RecordCount
End Function
Private Function myreplace() As Boolean
Dim aInsert(23, 1)
aInsert(0, 0) = "SER_NO"
aInsert(0, 1) = addstring(XSER_NO.text)

aInsert(1, 0) = "CHK_ID"
aInsert(1, 1) = addstring(XCHK_ID.text)

aInsert(2, 0) = "OLD"
aInsert(2, 1) = xOld.Value

aInsert(3, 0) = "code1"
aInsert(3, 1) = addstring(xCode1.text)

aInsert(4, 0) = "code2"
aInsert(4, 1) = addstring(xCode2.text)

aInsert(5, 0) = "[desca]"
aInsert(5, 1) = addstring(xdesca.text)

aInsert(6, 0) = "Bank_rec"
aInsert(6, 1) = addstring(xBANK_REC.text)

aInsert(7, 0) = "date_1"
aInsert(7, 1) = addDate(xDATE_1.text)

aInsert(8, 0) = "date_3"
aInsert(8, 1) = addDate(xDATE_3.text)

aInsert(9, 0) = "date_r"
aInsert(9, 1) = addDate(xDATE_R.text)

aInsert(10, 0) = "DateBank"
aInsert(10, 1) = addDate(xDateBank.text)

aInsert(11, 0) = "[VALUE]"
aInsert(11, 1) = Val(xValue.text)

aInsert(12, 0) = "NAME4"
aInsert(12, 1) = addstring(xNAME4.text)

aInsert(13, 0) = "BOX"
aInsert(13, 1) = addstring(XBOX.BoundText)

aInsert(14, 0) = "MEMO"
aInsert(14, 1) = addstring(xMEMO.text)

aInsert(15, 0) = "ID_BANK"
aInsert(15, 1) = addstring(XID_BANK.BoundText)

If Trim(xCode1.text) <> "" Then
    aInsert(16, 0) = "DESCA1"
    aInsert(16, 1) = addstring("⁄„Ì·:" & xNAME1.text)
ElseIf Trim(xCode2.text) <> "" Then
    aInsert(16, 0) = "DESCA1"
    aInsert(16, 1) = addstring("„Ê—œ:" & xName2.text)
ElseIf Trim(xacc.text) <> "" Then
    aInsert(16, 0) = "DESCA1"
    aInsert(16, 1) = addstring("Õ”«» :" & xacc.text)
End If

If xTransCode1.text <> "" Then
    aInsert(17, 0) = "descaTrans"
    aInsert(17, 1) = addstring("„Ê—œ: " & xTransName1.text)
Else
    aInsert(17, 0) = "descaTrans"
    aInsert(17, 1) = "NULL"
End If

If xTransCode2.text <> "" Then
    aInsert(17, 0) = "descaTrans"
    aInsert(17, 1) = addstring("⁄„Ì·: " & xTransName1.text)
Else
    aInsert(17, 0) = "descaTrans"
    aInsert(17, 1) = "NULL"
End If
 
aInsert(18, 0) = "Closed"
aInsert(18, 1) = addstring(retClose)

aInsert(19, 0) = "TransCode1"
aInsert(19, 1) = addstring(xTransCode1.text)

aInsert(20, 0) = "TransCode2"
aInsert(20, 1) = addstring(xTransCode2.text)

aInsert(21, 0) = "TransName1"
aInsert(21, 1) = addstring(xTransName1.text)

aInsert(22, 0) = "TransName2"
aInsert(22, 1) = addstring(xTransName2.text)

aInsert(23, 0) = "ACC"
aInsert(23, 1) = addstring(xacc.BoundText)

On Error GoTo myerror
con.BeginTrans
If XSER_NO.Enabled Then
    XSER_NO.text = RetZero(Val(Newflag("file5_20", "ser_no", con)), 6)
    aInsert(0, 1) = addstring(XSER_NO.text)
    con.Execute CreateInsert(aInsert, "file5_20")
Else
    con.Execute CreateUpdate(aInsert, "file5_20", " where SER_NO = " & addstring(XSER_NO.text))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub handleopt()
For i = 1 To 3
    optclose(i).Enabled = GetDesca("select ser_no from " & cFileName & " where closed = " & MyParn(i), con) <> ""
Next
End Sub
Private Sub xTransCode1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CLIENTLOOKUP 2
End Sub
Private Sub xTransCode1_LostFocus()
If Trim(xTransCode1.text) <> "" Then
    xTransName1.text = GetDesca("Select Desca from file4_10 where code = " & xTransCode1.text, con)
End If

If xTransCode1.text <> "" Then
    xTransCode2.text = ""
    xTransCode2.Enabled = False
    xTransName2.text = ""
Else
    xTransCode2.Enabled = True
End If
End Sub

Private Sub xTransCode2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CLIENTLOOKUP
End Sub
Private Sub xTransCode2_LostFocus()
If Trim(xTransCode2.text) <> "" Then
    xTransName2.text = GetDesca("Select Desca from file3_10 where code = " & xTransCode2.text, con)
    xTransCode1.text = ""
    xTransCode1.Enabled = False
    xTransName1.text = ""
Else
    xTransCode1.Enabled = True
End If
End Sub
Private Function retClose() As String
Dim i As Integer
For i = 0 To 2
    If xClosed(i).Value Then
        retClose = i & ""
        Exit For
    End If
Next
End Function
Private Sub xDateBank_GotFocus()
xDateBank.SelStart = 0
xDateBank.SelLength = Len(xDateBank.text)
End Sub
Private Sub xDescA_GotFocus()
xdesca.SelStart = 0
xdesca.SelLength = Len(xdesca.text)
End Sub
Private Sub xName2_GotFocus()
xName2.SelStart = 0
xName2.SelLength = Len(xName2.text)
End Sub
Private Sub XCode2_GotFocus()
xCode2.SelStart = 0
xCode2.SelLength = Len(xCode2.text)
End Sub
Private Sub xCode1_GotFocus()
xCode1.SelStart = 0
xCode1.SelLength = Len(xCode1.text)
End Sub
Private Sub xNAME1_GotFocus()
xNAME1.SelStart = 0
xNAME1.SelLength = Len(xNAME1.text)
End Sub
Private Sub xDATE_1_GotFocus()
xDATE_1.SelStart = 0
xDATE_1.SelLength = Len(xDATE_1.text)
End Sub
Private Sub xBANK_REC_GotFocus()
xBANK_REC.SelStart = 0
xBANK_REC.SelLength = Len(xBANK_REC.text)
End Sub
Private Sub xNAME4_GotFocus()
xNAME4.SelStart = 0
xNAME4.SelLength = Len(xNAME4.text)
End Sub
Private Sub xValue_GotFocus()
xValue.SelStart = 0
xValue.SelLength = Len(xValue.text)
End Sub
Private Sub XSER_NO_GotFocus()
XSER_NO.SelStart = 0
XSER_NO.SelLength = Len(XSER_NO.text)
End Sub
Private Sub XCHK_ID_GotFocus()
XCHK_ID.SelStart = 0
XCHK_ID.SelLength = Len(XCHK_ID.text)
End Sub
Private Sub xDATE_R_GotFocus()
xDATE_R.SelStart = 0
xDATE_R.SelLength = Len(xDATE_R.text)
End Sub
Private Sub xDATE_3_GotFocus()
xDATE_3.SelStart = 0
xDATE_3.SelLength = Len(xDATE_3.text)
End Sub
Private Sub xMEMO_GotFocus()
xMEMO.SelStart = 0
xMEMO.SelLength = Len(xMEMO.text)
End Sub
Private Sub xTransName2_GotFocus()
xTransName2.SelStart = 0
xTransName2.SelLength = Len(xTransName2.text)
End Sub
Private Sub xTransCode2_GotFocus()
xTransCode2.SelStart = 0
xTransCode2.SelLength = Len(xTransCode2.text)
End Sub
Private Sub xTransCode1_GotFocus()
xTransCode1.SelStart = 0
xTransCode1.SelLength = Len(xTransCode1.text)
End Sub
Private Sub xTransName1_GotFocus()
xTransName1.SelStart = 0
xTransName1.SelLength = Len(xTransName1.text)
End Sub
