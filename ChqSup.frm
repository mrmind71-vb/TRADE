VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form chqsupfrm 
   Caption         =   "√Ê—«ﬁ ﬁ»÷"
   ClientHeight    =   7530
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   9705
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
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin MSDataListLib.DataCombo XPART 
      Height          =   315
      Left            =   3735
      TabIndex        =   61
      Top             =   3870
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.TextBox XCHARGEDESCA 
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
      Left            =   1125
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   3465
      Width           =   5595
   End
   Begin VB.TextBox XCHARGE 
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
      Left            =   6795
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   3465
      Width           =   1230
   End
   Begin VB.Frame Frame6 
      Height          =   555
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   6435
      Width           =   3300
      Begin VB.CommandButton cmdNext 
         Height          =   375
         Left            =   2520
         Picture         =   "ChqSup.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "«· «·Ì"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   375
         Left            =   1755
         Picture         =   "ChqSup.frx":25C0
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         ToolTipText     =   "«·”«»ﬁ"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdLast 
         Height          =   375
         Left            =   810
         Picture         =   "ChqSup.frx":4B93
         Style           =   1  'Graphical
         TabIndex        =   51
         TabStop         =   0   'False
         ToolTipText     =   "√ŒÌ—"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   375
         Left            =   45
         Picture         =   "ChqSup.frx":726D
         Style           =   1  'Graphical
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   135
         Width           =   735
      End
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
      Height          =   690
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
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   855
      RightToLeft     =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   675
      Width           =   1530
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
      Left            =   2745
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   19
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
      Left            =   6795
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
      ScaleWidth      =   9705
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   0
      Width           =   9705
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   2610
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSup.frx":994C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   6030
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSup.frx":9A96
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Õ›Ÿ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdDel 
         Height          =   420
         Left            =   4890
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSup.frx":9ED8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   420
         Left            =   3750
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSup.frx":C772
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdAdd 
         Height          =   420
         Left            =   7170
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSup.frx":ECEB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   8310
         Picture         =   "ChqSup.frx":11297
         Style           =   1  'Graphical
         TabIndex        =   43
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
         Left            =   180
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSup.frx":13A6A
         RightToLeft     =   -1  'True
         TabIndex        =   42
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
      ScaleWidth      =   9705
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7080
      Width           =   9705
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
         Left            =   2325
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   75
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
         Left            =   6210
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   90
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
         Left            =   4995
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   90
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
         Left            =   3600
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   90
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
      Left            =   6795
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
      Left            =   2745
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   18
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
      Left            =   5310
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
      Left            =   5310
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
      Left            =   5670
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
      Left            =   5670
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
      Left            =   5310
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2385
      Width           =   2715
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   90
      Top             =   1350
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
   Begin MSDataListLib.DataCombo XID_BANK 
      Height          =   315
      Left            =   5310
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
      Left            =   3825
      Top             =   1980
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
      Left            =   90
      Top             =   1665
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
   Begin VB.Frame Frame1 
      Height          =   1410
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   4950
      Width           =   7845
      Begin VB.TextBox xDATE_3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3300
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   14
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
         TabIndex        =   16
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   150
         Width           =   1740
      End
      Begin MSDataListLib.DataCombo XBOX 
         Height          =   315
         Left            =   225
         TabIndex        =   15
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
         TabIndex        =   37
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ«  :"
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
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1050
         Width           =   675
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «·”œ«œ/  ŸÂÌ— / —›÷"
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
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   675
         Width           =   1860
      End
   End
   Begin MSDataListLib.DataCombo xMosm 
      Height          =   315
      Left            =   2925
      TabIndex        =   54
      Top             =   585
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   90
      Top             =   1035
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
   Begin MSDataListLib.DataCombo XBRANCH 
      Height          =   315
      Left            =   2925
      TabIndex        =   59
      Top             =   945
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
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
   Begin MSDataListLib.DataCombo xacc 
      Height          =   315
      Left            =   3735
      TabIndex        =   63
      Top             =   4275
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSAdodcLib.Adodc data7 
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
      Left            =   8145
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   4320
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·‘—Ìﬂ : "
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
      Left            =   8145
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   3915
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
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
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   945
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„’—Ê›  :"
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
      Left            =   8145
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   3510
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„Ê”„"
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
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   630
      Width           =   495
   End
   Begin VB.Label LabelCode2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   1755
      Width           =   915
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
      TabIndex        =   32
      Top             =   75
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "»‰ﬂ «·”Õ» :"
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
      Index           =   0
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   3150
      Width           =   1005
   End
   Begin VB.Label label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3465
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   2745
      Width           =   525
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   2070
      Width           =   795
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3465
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   2430
      Width           =   1680
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   2115
      Width           =   1380
   End
   Begin VB.Label LabelCode 
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
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1440
      Width           =   885
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   2835
      Width           =   570
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2430
      Width           =   945
   End
End
Attribute VB_Name = "chqsupfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim bSumMode As Boolean
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
CmdAdd.Enabled = (nMode = LoadMode And optclose(0).Value) And bedit
cmdSave.Enabled = (nMode = LoadMode Or optclose(0).Value) And bedit
cmdSum.Enabled = (nMode = LoadMode Or optclose(0).Value) And bedit
CmdUndo.Enabled = (nMode = LoadMode Or optclose(0).Value)
CmdDel.Enabled = (nMode = LoadMode) And bedit
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
XSER_NO.Enabled = Not (nMode = LoadMode)
End Sub
Sub CardLookup()
Dim Generalarray(5)
Dim listarray(2, 7)
Dim GrdArray(8, 1)

If optclose(0).Value Then cWhere = "CLOSED = '0'"
If optclose(1).Value Then cWhere = "CLOSED = '1'"
If optclose(2).Value Then cWhere = "CLOSED = '2'"

Set Generalarray(0) = Me
Generalarray(1) = " SELECT FILE5_21.SER_NO,FILE5_21.DESCA1,CONVERT(VARCHAR(10),FILE5_21.DATE_1 ,111), CONVERT(VARCHAR(10),FILE5_21.DATE_r ,111), [VALUE] , CHK_ID , FILE8_51.DESCA , FILE5_21.DESCA AS CHQDESCA , FILE5_21.ACC From FILE5_21 LEFT JOIN FILE8_51 ON FILE8_51.CODE = FILE5_21.CHARGE " & turn(cWhere, " where ") & cWhere
Generalarray(2) = ""
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "„”·”· «Ê „” ›Ìœ -  «—ÌŒ  Õ—Ì—-—ﬁ„ «·‘Ìﬂ- „’—Ê›-»Ì«‰"
listarray(0, 1) = "(%%FILE8_51.desca%% or %%FILE5_21.SER_NO%% or %%FILE5_21.desca1%% or %%FILE5_21.desca%% or chk_id Like '%cFilter%' Or ##Date_R##)"

listarray(1, 0) = " «—ÌŒ «” Õﬁ«ﬁ"
listarray(1, 1) = " ##Date_1##"


listarray(2, 0) = "«·ﬁÌ„…"
listarray(2, 1) = "**[value]**"

GrdArray(0, 0) = "„”·”·"
GrdArray(0, 1) = 800

GrdArray(1, 0) = "‘Ìﬂ „‰"
GrdArray(1, 1) = 2000

GrdArray(2, 0) = "«” Õﬁ«ﬁ"
GrdArray(2, 1) = 1300

GrdArray(3, 0) = " Õ—Ì—"
GrdArray(3, 1) = 1300

GrdArray(4, 0) = "ﬁÌ„…"
GrdArray(4, 1) = 1200

GrdArray(5, 0) = "—ﬁ„ «·‘Ìﬂ"
GrdArray(5, 1) = 1400

GrdArray(6, 0) = "«·„’—Ê›"
GrdArray(6, 1) = 2000

GrdArray(7, 0) = "«·»Ì«‰"
GrdArray(7, 1) = 2000

GrdArray(8, 0) = "Õ”«»"
GrdArray(8, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "«” ⁄·«„ "
Search3.Show 1
End Sub
Sub CLIENTLOOKUP(Optional nFlag As Integer = 1)

Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
If nFlag = 1 Then
    Generalarray(1) = "SELECT Code , Desca  From file4_10"
Else
    Generalarray(1) = "SELECT Code , Desca  From file3_10"
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
Search3.Caption = "«” ⁄·«„" & " " & IIf(nFlag = 1, "«·„Ê—œÌ‰", "«·⁄„·«¡")
Search3.Show 1
End Sub
Sub myDefine()

XSER_NO.text = RetZero(Val(Newflag("FILE5_21", "ser_no", con)), 6)
xCharge.text = ""
XCHARGEDESCA.text = ""
XBRANCH.BoundText = ""
xacc.BoundText = ""
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
xBox.BoundText = ""
xPart.BoundText = ""
xValue.text = ""
xMEMO.text = ""
xMosm.BoundText = ""
xDesca.text = ""
xClosed(0) = True
xClosed(1) = False
xClosed(2) = False
Handlecontrols DefineMode
End Sub
Sub myDefine2()

XSER_NO.text = RetZero(Val(Newflag("FILE5_21", "ser_no", con)), 6)
XCHK_ID.text = Val(XCHK_ID.text) + 1
xOld.Value = 0
xName2.text = ""
xNAME4.text = ""
xDATE_1.text = ""
xValue.text = ""
xMEMO.text = ""
xDesca.text = ""
xClosed(0) = True
xClosed(1) = False
xClosed(2) = False
xacc.BoundText = ""
xPart.BoundText = ""
xCharge.text = ""

Handlecontrols DefineMode
End Sub
Sub myProc()
If TypeOf ActiveControl Is TextBox Then
    ActiveControl.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
Else
    CardTable.Find "SER_NO = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload
    Unload Search3
End If
End Sub
Sub myload()
XSER_NO.text = CardTable!Ser_no
xCharge.text = CardTable!CHARGE & ""
If Not IsNull(CardTable!CHARGE) Then
    XCHARGEDESCA.text = GetDesca("SELECT DESCA FROM FILE8_51 WHERE CODE = " & MyParn(CardTable!CHARGE), con) & ""
Else
    XCHARGEDESCA.text = ""
End If

XCHK_ID.text = CardTable!CHK_ID & ""
xCode1.text = CardTable!Code1 & ""
xCode2.text = CardTable!CODE2 & ""
xNAME1.text = CardTable!desca1 & ""
xName2.text = CardTable!Desca2 & ""
xNAME4.text = CardTable!NAME4 & ""
xDesca.text = CardTable!DESCA & ""
xClosed(0).Value = IIf(CardTable!CLOSED = "0", True, False)
xClosed(1).Value = IIf(CardTable!CLOSED = "1", True, False)
xClosed(2).Value = IIf(CardTable!CLOSED = "2", True, False)

xOld.Value = IIf(CardTable!OLD, 1, 0)
xBANK_REC.text = CardTable!Bank_rec & ""
xValue.text = Format(CardTable!Value, "Fixed")
xDATE_1.text = Format(CardTable!date_1, "dd-mm-yyyy")
xDATE_3.text = Format(CardTable!date_3, "dd-mm-yyyy")
xDATE_R.text = Format(CardTable!date_R, "dd-mm-yyyy")
xMEMO.text = CardTable!Memo & ""
XID_BANK.BoundText = CardTable!ID_BANK & ""
xBox.BoundText = CardTable!BOX & ""
xPart.BoundText = CardTable!PART & ""
xMosm.BoundText = CardTable!MOSM & ""
XBRANCH.BoundText = CardTable!branch & ""
xacc.BoundText = CardTable!ACC & ""

Handlecontrols LoadMode
End Sub
Function MYVALID() As Boolean
If XSER_NO.text = "" Then
    MsgBox "ÌÃ»  ”ÃÌ· „””·”· ··‘Ìﬂ"
    Exit Function
End If
If xMosm.BoundText = "" Then
    MsgBox "ÌÃ»  ”ÃÌ· «·„Ê”„ "
    Exit Function
End If

If Not IsDate(xDATE_R.text) Then
    MsgBox "ÌÃ»  ”ÃÌ·  «—ÌŒ «· Õ—Ì—"
    Exit Function
End If

If xCode1.text <> "" Then
    If GetDesca("select code from file4_10 where code = " & MyParn(xCode1.text), con) = "" Then Exit Function
End If

If xCode2.text <> "" Then
    If GetDesca("select code from file3_10 where code = " & MyParn(xCode2.text), con) = "" Then Exit Function
End If


If XSER_NO.Enabled And Trim(xCode1.text) <> "" Then
    cString = GetDesca("select ser_no from FILE5_21 where code1 = " & MyParn(xCode1.text) & " and Chk_Id = " & MyParn(XCHK_ID.text), con) & ""
    If Trim(cString) <> "" Then
        MsgBox "‘Ìﬂ »‰›” «·—ﬁ„ ·‰›” «·≈”„"
        Exit Function
    End If
End If

If xCode1.text = "" And xCode2.text = "" And xCharge.text = "" And xPart.BoundText = "" And xacc.BoundText = "" Then
    MsgBox "ÌÃ»  ÕœÌœ «·„” ›Ìœ "
    Exit Function
End If

MYVALID = True
End Function
Private Sub CmdAdd_Click()
myDefine
On Error Resume Next
XCHK_ID.SetFocus
Err.Clear
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
con.BeginTrans
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    con.Execute "delete  From FILE5_21 where Ser_No = " & MyParn(XSER_NO.text)
End If
con.CommitTrans

CardTable.Requery
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.Find "SER_NO < " & MyParn(XSER_NO.text), , adSearchBackward, adBookmarkLast
    If CardTable.EOF Then CardTable.MoveFirst
    myload
Else
    If optclose(0).Value Then CmdAdd_Click Else myDefine
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdExit_Click()
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
        myDefine
    End If
End If
End Sub
Private Sub CmdUndo_Click()
bSumMode = False
If CardTable.EOF And CardTable.BOF Then
    myDefine
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

openCon con
data1.ConnectionString = strCon
data1.RecordSource = "FILE0_50"

Set xBox.RowSource = data1
xBox.ListField = "Desca"
xBox.BoundColumn = "Code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "FILE5_10"

Set XID_BANK.RowSource = DATA2
XID_BANK.ListField = "Desca"
XID_BANK.BoundColumn = "code"

data4.ConnectionString = strCon
data4.RecordSource = "SELECT * FROM MOSM"
Set xMosm.RowSource = data4
xMosm.ListField = "DESCA"
xMosm.BoundColumn = "MOSM"
xMosm = cPMosm

data5.ConnectionString = strCon
data5.RecordSource = "SELECT * FROM BRANCH "
Set XBRANCH.RowSource = data5
XBRANCH.ListField = "DESCA"
XBRANCH.BoundColumn = "CODE"

DATA6.ConnectionString = strCon
DATA6.RecordSource = "SELECT * FROM FILE8_71 ORDER BY DESCA"
Set xPart.RowSource = DATA6
xPart.ListField = "DESCA"
xPart.BoundColumn = "CODE"

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
closeCon con
End Sub
Private Sub optclose_Click(Index As Integer)
cString = "Select FILE5_21.*,FILE4_10.DESCA AS DESCA1,FILE3_10.DESCA AS DESCA2" & _
          " From (FILE5_21 left JOIN file4_10 ON FILE5_21.CODE1 = FILE4_10.CODE) LEFT JOIN FILE3_10 ON FILE5_21.CODE2 = FILE3_10.CODE"
If Index <> 3 Then
    cString = cString & turn(cString) & " closed = " & MyParn(Index)
End If
cString = cString & " Order by Ser_No"

Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenKeyset, adLockOptimistic, adCmdText

If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    If optclose(0).Value Then
        CmdAdd_Click
    Else
        myDefine
    End If
End If
End Sub
Private Sub XBOX_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xBox.BoundText = ""
End Sub

Private Sub xBOX2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xBox2.BoundText = ""
End Sub

Private Sub XCHK_ID_Validate(Cancel As Boolean)
If publicFlag = 2 And XSER_NO.Enabled Then
    If GetDesca("SELECT SER_NO FROM FILE5_21 WHERE SER_NO <> " & MyParn(XSER_NO.text) & " AND  CHK_ID = " & MyParn(XCHK_ID.text), con) <> "" Then
        MsgBox "—ﬁ„ «·‘Ìﬂ „ﬂ—— „‰ ﬁ»·"
    End If
End If
End Sub
Private Sub xClosed_Click(Index As Integer)
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
xNAME1.text = GetDesca("Select Desca from FILE4_10 where code = " & MyParn(xCode1.text), con)
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
If Trim(xCode2.text) = "" Then Exit Sub
xName2.text = ""
xCode2.text = RetZero(xCode2.text, 3)
xName2.text = GetDesca("Select Desca from FILE3_10 where code = " & MyParn(xCode2.text), con)

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
Private Function myreplace() As Boolean
Dim aInsert(21, 1)
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
aInsert(5, 1) = addstring(xDesca.text)

aInsert(6, 0) = "Bank_rec"
aInsert(6, 1) = addstring(xBANK_REC.text)

aInsert(7, 0) = "date_1"
aInsert(7, 1) = addDate(xDATE_1.text)

aInsert(8, 0) = "date_3"
aInsert(8, 1) = addDate(xDATE_3.text)

aInsert(9, 0) = "date_r"
aInsert(9, 1) = addDate(xDATE_R.text)

aInsert(10, 0) = "[VALUE]"
aInsert(10, 1) = Val(xValue.text)

aInsert(11, 0) = "NAME4"
aInsert(11, 1) = addstring(xNAME4.text)

aInsert(12, 0) = "BOX"
aInsert(12, 1) = addstring(xBox.BoundText)

aInsert(13, 0) = "MEMO"
aInsert(13, 1) = addstring(xMEMO.text)

aInsert(14, 0) = "ID_BANK"
aInsert(14, 1) = addstring(XID_BANK.BoundText)

aInsert(18, 0) = "CHARGE"
aInsert(18, 1) = addstring(xCharge.text)


If Trim(xCode1.text) <> "" Then
    aInsert(15, 0) = "DESCA1"
    aInsert(15, 1) = addstring("„Ê—œ:" & xNAME1.text)
ElseIf Trim(xCode2.text) <> "" Then
    aInsert(15, 0) = "DESCA1"
    aInsert(15, 1) = addstring("⁄„Ì·:" & xNAME1.text)
Else
    aInsert(15, 0) = "DESCA1"
    aInsert(15, 1) = "NULL"
End If


aInsert(16, 0) = "Closed"
aInsert(16, 1) = addstring(retClose)

aInsert(17, 0) = "MOSM"
aInsert(17, 1) = addstring(xMosm.BoundText)

aInsert(19, 0) = "BRANCH"
aInsert(19, 1) = addstring(XBRANCH.BoundText)

aInsert(20, 0) = "PART"
aInsert(20, 1) = addstring(xPart.BoundText)

aInsert(21, 0) = "ACC"
aInsert(21, 1) = addstring(xacc.BoundText)

On Error GoTo myerror
con.BeginTrans
If XSER_NO.Enabled Then
    XSER_NO.text = RetZero(Val(Newflag("FILE5_21", "ser_no", con)), 6)
    aInsert(0, 1) = addstring(XSER_NO.text)
    con.Execute CreateInsert(aInsert, "FILE5_21")
Else
    con.Execute CreateUpdate(aInsert, "FILE5_21", " where SER_NO = " & addstring(XSER_NO.text))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Function retClose() As String
Dim i As Integer
For i = 0 To 2
    If xClosed(i).Value Then
        retClose = i & ""
        Exit For
    End If
Next
End Function
Private Sub xDescA_GotFocus()
xDesca.SelStart = 0
xDesca.SelLength = Len(xDesca.text)
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
Private Sub XCHARGE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then ChargeLookup
End Sub

Private Sub XCHARGE_LostFocus()
    If xCharge.text <> "" Then
        XCHARGEDESCA.text = GetDesca("SELECT DESCA FROM FILE8_51 WHERE CODE = " & MyParn(xCharge.text), con)
    Else
        XCHARGEDESCA.text = ""
    End If
End Sub

Private Sub ChargeLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FILE8_51 "
Generalarray(2) = "Order by code"
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
End Sub


