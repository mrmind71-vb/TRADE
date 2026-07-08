VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Clients_sub 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "⁄„·«¡ ‰ﬁœÏ œ«∆„Ì‰ ··„Õ·« "
   ClientHeight    =   6315
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   8775
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
   ScaleHeight     =   6315
   ScaleWidth      =   8775
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   990
      Top             =   2070
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
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   8775
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5820
      Width           =   8775
      Begin VB.CommandButton FIX 
         Caption         =   "FIX"
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   90
         Width           =   1050
      End
      Begin VB.CommandButton cmdNext 
         Height          =   375
         Left            =   4605
         Picture         =   "Clients_sub.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "«· «·Ì"
         Top             =   90
         Width           =   915
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   375
         Left            =   3690
         Picture         =   "Clients_sub.frx":25C0
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "«·”«»ﬁ"
         Top             =   90
         Width           =   915
      End
      Begin VB.CommandButton cmdLast 
         Height          =   375
         Left            =   2535
         Picture         =   "Clients_sub.frx":4B93
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "√ŒÌ—"
         Top             =   90
         Width           =   915
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   375
         Left            =   1665
         Picture         =   "Clients_sub.frx":726D
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   90
         Width           =   870
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   8775
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   1890
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Clients_sub.frx":994C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   5310
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Clients_sub.frx":9A96
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Õ›Ÿ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdDel 
         Height          =   420
         Left            =   4170
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Clients_sub.frx":9ED8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   420
         Left            =   3030
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Clients_sub.frx":C772
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdAdd 
         Height          =   420
         Left            =   6450
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Clients_sub.frx":ECEB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   7590
         Picture         =   "Clients_sub.frx":11297
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   45
         Width           =   1140
      End
      Begin Threed.SSCommand CMD_PrintBar 
         Height          =   420
         Left            =   90
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   45
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   741
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
         Picture         =   "Clients_sub.frx":13A6A
         Caption         =   "ÿ»«⁄… »«—ﬂÊœ"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   900
      Top             =   1890
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
      Height          =   1365
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   540
      Width           =   8385
      Begin VB.CheckBox xIsOk 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ«—  „›⁄·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox XBARCODE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   15
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   135
         Width           =   1725
      End
      Begin VB.TextBox XCODE2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   135
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   900
         Width           =   1050
      End
      Begin VB.TextBox xDescA 
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
         Left            =   1440
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   5145
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   5220
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1365
      End
      Begin VB.TextBox xManager 
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
         Left            =   1440
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   5145
      End
      Begin Threed.SSCommand BARCODE 
         Height          =   420
         Left            =   315
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   135
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   741
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
         Picture         =   "Clients_sub.frx":163EC
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·»ÿ«ﬁ… :"
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
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   990
         Width           =   990
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ :"
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
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«”„ :"
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
         Left            =   6705
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   570
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2310
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1890
      Width           =   8430
      Begin VB.TextBox xB_DATE 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   135
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1845
         Width           =   2175
      End
      Begin VB.TextBox xEMAIL 
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
         Left            =   135
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1485
         Width           =   6450
      End
      Begin VB.TextBox xFAx 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   135
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1125
         Width           =   2220
      End
      Begin VB.TextBox xPhone1 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4590
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1125
         Width           =   1995
      End
      Begin VB.TextBox xAddress 
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
         Height          =   555
         Left            =   90
         MaxLength       =   100
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   547
         Width           =   6495
      End
      Begin MSDataListLib.DataCombo xgroup 
         Height          =   315
         Left            =   3510
         TabIndex        =   12
         Top             =   180
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label xF_date 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1845
         Width           =   1905
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·„Ì·«œ :"
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
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1935
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "»œ«Ì… ’·«ÕÌ… «·ﬂ«—  :"
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
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1935
         Width           =   1620
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·≈Ì„Ì· :"
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
         Left            =   6705
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1575
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "„Ê»Ì· :"
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
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1170
         Width           =   540
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«· ·Ì›Ê‰ :"
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
         Left            =   6705
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1170
         Width           =   705
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ã„Ê⁄… :"
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
         Left            =   6705
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   225
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄‰Ê«‰ :"
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
         Left            =   6705
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   540
         Width           =   645
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "„·«ÕŸ« "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   4230
      Width           =   8430
      Begin VB.TextBox xRemark 
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
         Height          =   1140
         Left            =   45
         MaxLength       =   200
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   270
         Width           =   8295
      End
   End
End
Attribute VB_Name = "Clients_sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CON As New ADODB.Connection
Public myFlag As Integer
Dim formMode As Byte
Dim CardTable As New ADODB.Recordset
Const LoadMode = 1, DefineMode = 2

Private Sub BARCODE_Click()
    If xIsOk.Value = 0 Then Exit Sub
    CON.BeginTrans
    On Error GoTo myerror
    cString = "INSERT INTO ADDPRINT3(barcode,desca ,ISPRINT)"
    cString = cString & _
                "Values(" & _
                addstring(XBARCODE.text) & "," & _
                addstring(xDescA.text) & "," & _
                1 & _
                       ")"
    CON.Execute cString
    CON.CommitTrans
    Inform " „ «· ÕÊÌ· »‰Ã«Õ"
    Exit Sub
myerror:
    CON.RollbackTrans
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub CMD_PrintBar_Click()
    BarCodeCust.Show 1
End Sub

Private Sub fix_Click()
    Dim cCode As String
    Dim cCodeSEC As String
    If InputBox("", "") <> "MOR2013" Then Exit Sub
'   con.Execute " UPDATE FILE3_20 SET CODE2 = CODE "
    cCode = "A00000"
    CardTable.MoveFirst
    Do While Not CardTable.EOF
        cCode = CardTable!CODE
        cCodeSEC = RetRand(1) & Mid(cCode, 2)
        CON.Execute " UPDATE FILE3_20  SET BARCODE = " & addstring(cCodeSEC) & " WHERE CODE = " & MyParn(CardTable!CODE)
        CardTable.MoveNext
        Me.Caption = cCode
    Loop
    MsgBox " O K "
End Sub
Private Sub Form_Load()
CON.CursorLocation = adUseClient
CON.Open strCon

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM file3_51"
Set xGroup.RowSource = data1
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

openCardTable
myUndo
End Sub
Private Sub CmdAdd_Click()
mydefine
xCode.SetFocus
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    CON.BeginTrans
    CON.Execute "Delete  From FILE3_20 Where code = " & MyParn(xCode.text)
    CON.CommitTrans
    openCardTable
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "code < " & MyParn(xCode.text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    Else
        mydefine
    End If
End If
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
    CON.RollbackTrans
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
openCardTable
myUndo
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CmdInform_Click()
    CustSUPPLookupAll Me, Search3
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
Sub Handlecontrols(nMode)
CmdAdd.Enabled = (nMode = LoadMode)
CmdDel.Enabled = (nMode = LoadMode)
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
xCode.Enabled = Not (nMode = LoadMode)
End Sub
Sub mydefine()
Dim cCode As String
xCode.text = IncRec(GetDesca("SELECT MAX(CODE) FROM FILE3_20"), CON)
XBARCODE.text = RetRand(1) & Mid(xCode.text, 2)
xCode2.text = ""
xIsOk.Value = False
xDescA.text = ""
xManager.text = ""
xEMAIL.text = ""
xGroup.BoundText = ""

xAddress.text = ""
xPhone1.text = ""
xFAx.text = ""

xF_date.Caption = Format(Date, "DD-MM-YYYY")
xB_DATE.text = ""


xRemark.text = ""
Handlecontrols DefineMode
End Sub
Sub myload()
xCode.text = CardTable!CODE & ""
XBARCODE.text = CardTable!BARCODE & ""
xCode2.text = CardTable!CODE2 & ""
xIsOk.Value = IIf(CardTable!ISOK, 1, 0)
xDescA.text = CardTable!DESCA
xManager.text = CardTable!Manager & ""
xAddress.text = CardTable!Address & ""
xEMAIL.text = CardTable!email & ""
xPhone1.text = CardTable!PHONE1 & ""
xFAx.text = CardTable!Fax & ""
xF_date.Caption = Format(CardTable!F_DATE, "dd-mm-yyyy")
xB_DATE.text = Format(CardTable!F_DATE, "dd-mm-yyyy")
xGroup.BoundText = CardTable!Group & ""
xRemark.text = CardTable!remark & ""
xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim nTry As Integer
Dim aInsert(14, 1)
XBARCODE.text = RetRand(1) & Mid(xCode.text, 2)
aInsert(0, 0) = "Code"
aInsert(0, 1) = addstring(xCode.text)

aInsert(1, 0) = "desca"
aInsert(1, 1) = addstring(xDescA.text)

aInsert(2, 0) = "Manager"
aInsert(2, 1) = addstring(xManager.text)

aInsert(3, 0) = "Address"
aInsert(3, 1) = addstring(xAddress.text)

aInsert(4, 0) = "phone1"
aInsert(4, 1) = addstring(xPhone1.text)

aInsert(5, 0) = "fax"
aInsert(5, 1) = addstring(xFAx.text)

aInsert(6, 0) = "[Group]"
aInsert(6, 1) = addstring(xGroup.BoundText)

aInsert(8, 0) = "f_date"
aInsert(8, 1) = addDate(xF_date.Caption)

aInsert(9, 0) = "b_date"
aInsert(9, 1) = addDate(xB_DATE.text)

aInsert(10, 0) = "email"
aInsert(10, 1) = addstring(xEMAIL.text)

aInsert(12, 0) = "remark"
aInsert(12, 1) = addstring(xRemark.text)

aInsert(13, 0) = "BARCODE"
aInsert(13, 1) = addstring(XBARCODE.text)

aInsert(14, 0) = "IsOk"
aInsert(14, 1) = xIsOk.Value

CON.BeginTrans
If xCode.Enabled Then
    aInsert(0, 1) = addstring(xCode.text)
    cString = CreateInsert(aInsert, "file3_20")
    CON.Execute cString
Else
    cString = CreateUpdate(aInsert, "file3_20", " WHERE CODE = " & MyParn(xCode.text))
    CON.Execute cString
End If
CON.CommitTrans
myreplace = True
Exit Function
myerror:
CON.RollbackTrans
If Err.Number <> 0 Then MsgBox Err.Description
Err.Clear
End Function
Sub myProc()
   CardTable.Find "CODE = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
   myload
   Search3.Hide
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon CON
Err.Clear
End Sub

Private Sub xCode_LostFocus()
If xCode.text = "" Then Exit Sub
xCode.text = xCode.text
CardTable.Find "CODE = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Function MYVALID() As Boolean
If xCode.text = "" Then
    MsgBox "«·ﬂÊœ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If
If xDescA.text = "" Then
    MsgBox "«·≈”„ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If
MYVALID = True
End Function
Private Sub myUndo()
'On Error GoTo myerror
If CardTable.BOF And CardTable.EOF Then
    mydefine
Else
    If xCode.text <> "" Then
        CardTable.Find "CODE = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "Select * From file3_20"
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY CODE"
CardTable.Open cString, CON, adOpenStatic, adLockReadOnly, adCmdText
End Sub

