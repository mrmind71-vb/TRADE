VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Clients 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  ⁄„·«¡ √Ã·"
   ClientHeight    =   6975
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   9645
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   9645
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
      Height          =   735
      Left            =   1215
      MaxLength       =   200
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Top             =   5670
      Width           =   6630
   End
   Begin VB.Frame Frame1 
      Height          =   1680
      Left            =   855
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   540
      Width           =   8520
      Begin VB.CheckBox xIsMonth 
         Alignment       =   1  'Right Justify
         Caption         =   " Ã„Ì⁄ ‘Â—Ï ··›Ê« Ì—"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   540
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   180
         Width           =   1590
      End
      Begin VB.CheckBox xIsOnest 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄„Ì· «„«‰« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   225
         Width           =   1275
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
         Left            =   1890
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   540
         Width           =   4695
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
         Height          =   360
         Left            =   5265
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   150
         Width           =   1320
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
         Left            =   1890
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   900
         Width           =   4695
      End
      Begin MSDataListLib.DataCombo xAcc 
         Height          =   315
         Left            =   2205
         TabIndex        =   42
         Top             =   1260
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Õ”«» : "
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
         Left            =   6750
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1305
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "≈”„ «·‘—ﬂ… : "
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
         Left            =   6750
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   990
         Width           =   1140
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
         Left            =   6795
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
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
         Left            =   6795
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   615
         Width           =   570
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   720
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
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   855
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   4635
      Width           =   8520
      Begin VB.TextBox xf_Balance 
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
         Left            =   225
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   2130
      End
      Begin VB.TextBox xf_Date 
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
         Left            =   5040
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox xDisc 
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
         Left            =   5040
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   540
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ √Ê · «·„œ… :"
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
         TabIndex        =   8
         Top             =   270
         Width           =   1290
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "—’Ìœ √Ê· «·„œ… :"
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
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… «·Œ’„ :"
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
         Top             =   645
         Width           =   1080
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   9645
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6510
      Width           =   9645
      Begin VB.CommandButton cmdFirst 
         Height          =   375
         Left            =   2835
         Picture         =   "Clients.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdLast 
         Height          =   375
         Left            =   3705
         Picture         =   "Clients.frx":26DF
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "√ŒÌ—"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   375
         Left            =   4860
         Picture         =   "Clients.frx":4DB9
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "«·”«»ﬁ"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdNext 
         Height          =   375
         Left            =   5775
         Picture         =   "Clients.frx":738C
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "«· «·Ì"
         Top             =   45
         Width           =   915
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
      ScaleWidth      =   9645
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9645
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   6915
         Picture         =   "Clients.frx":994C
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   45
         Width           =   1140
      End
      Begin VB.CommandButton CmdAdd 
         Height          =   420
         Left            =   5739
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Clients.frx":C11F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   420
         Left            =   2211
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Clients.frx":E6CB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdDel 
         Height          =   420
         Left            =   3387
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Clients.frx":10C44
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   4563
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Clients.frx":134DE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Õ›Ÿ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   1035
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Clients.frx":13920
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2355
      Left            =   855
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2250
      Width           =   8520
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   330
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   180
         Width           =   330
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
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1980
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
         Height          =   360
         Left            =   135
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1562
         Width           =   6450
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
         Height          =   360
         Left            =   135
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1146
         Width           =   6450
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
         TabIndex        =   17
         Top             =   535
         Width           =   6495
      End
      Begin MSDataListLib.DataCombo xgroup 
         Height          =   315
         Left            =   2205
         TabIndex        =   19
         Top             =   180
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «· ”ÃÌ· «·÷—Ì»Ï"
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
         TabIndex        =   26
         Top             =   1980
         Width           =   1740
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
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1620
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   18
         Top             =   540
         Width           =   645
      End
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   5940
      Width           =   780
   End
End
Attribute VB_Name = "Clients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public myFlag As Integer
Dim formMode As Byte, cTableName As String, cGroupname As String
Dim CardTable As New ADODB.Recordset
Const LoadMode = 1, DefineMode = 2
Private Sub Command3_Click()
ReDim aLocal(6)
aLocal(0) = "FILE3_50"
aLocal(1) = "Code"
aLocal(2) = "Desca"
aLocal(3) = "ﬂÊœ „Ã„Ê⁄…"
aLocal(4) = "»Ì«‰ „Ã„Ê⁄« "
aLocal(5) = "„Ã„Ê⁄«  ⁄„·«¡"
aLocal(6) = 2
FlagFrm2.bedit = True
FlagFrm2.myPublic = aLocal
FlagFrm2.Show 1
data1.Refresh
End Sub
Private Sub fix_Click()
End Sub

Private Sub Form_Load()
con.CursorLocation = adUseClient
con.Open strCon

Me.Caption = "»Ì«‰«  «·⁄„·«¡"


Set data1.Recordset = myRecordSet("SELECT * FROM file3_50", con)
Set xGroup.RowSource = data1
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

Set DATA2.Recordset = myRecordSet("SELECT * FROM ACC0_10 ORDER BY CODE ", con)
Set xacc.RowSource = DATA2
xacc.ListField = "Desca"
xacc.BoundColumn = "Code"

CardTable.CursorLocation = adUseClient
CardTable.Open "Select * From file3_10 order by code", con, adOpenKeyset, adLockReadOnly, adCmdText
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    myDefine
End If
End Sub
Private Sub CmdAdd_Click()
myDefine
xCode.SetFocus
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    con.Execute "Delete  From FILE3_10 Where code = " & MyParn(xCode.text)
    AddLod_Data cusername, 2, " Õ–› ⁄„·«¡ ", con, xCode.text, , , xDesca.text
    
    CardTable.Requery
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "code < " & MyParn(xCode.text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    Else
        myDefine
    End If
End If
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
AddLod_Data cusername, 1, " Õ÷Ÿ ⁄„·«¡ ", con, xCode.text, , , xDesca.text

CardTable.Requery
CardTable.Find "code = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.EOF And CardTable.BOF Then
    myDefine
Else
    If xCode.Enabled Then
        CardTable.MoveLast
    Else
        CardTable.Find "code = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CmdInform_Click()
    CustLookupAll Me, Search3
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
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FILE3_10"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = True

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search
Search.Caption = "≈” ⁄·«„ "
Search.Show 1
End Sub
Sub myDefine()


xCode.text = RetZero(Newflag("file3_10", "code", con), 4)
xDesca.text = ""
xManager.text = ""
XEMAIL.text = ""
xGroup.BoundText = ""
xacc.BoundText = ""
XISONEST.Value = 0
xIsMonth.Value = 0
xAddress.text = ""
xPhone1.text = ""
xFAx.text = ""

xF_date.text = ""
XF_BALANCE.text = ""
XDISC.text = ""

xRemark.text = ""
Handlecontrols DefineMode
End Sub
Sub myload()
xCode.text = CardTable!CODE & ""
xDesca.text = CardTable!Desca
xManager.text = CardTable!Manager & ""
xAddress.text = CardTable!Address & ""
XEMAIL.text = CardTable!email & ""
xPhone1.text = CardTable!PHONE1 & ""
xFAx.text = CardTable!Fax & ""
XISONEST.Value = IIf(CardTable!ISONEST, 1, 0)
xIsMonth.Value = IIf(CardTable!IsMonth, 1, 0)
xF_date.text = Format(CardTable!F_DATE, "dd-mm-yyyy")
XF_BALANCE.text = Format(CardTable!F_Balance, "##0.00")
XDISC.text = Format(CardTable!disc, "##0.00")
xGroup.BoundText = CardTable!Group & ""
xacc.BoundText = CardTable!ACC & ""
xRemark.text = CardTable!remark & ""
xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim nTry As Integer
Dim aInsert(14, 1)
aInsert(0, 0) = "Code"
aInsert(0, 1) = addstring(xCode.text)

aInsert(1, 0) = "desca"
aInsert(1, 1) = addstring(xDesca.text)

aInsert(2, 0) = "Manager"
aInsert(2, 1) = addstring(xManager.text)

aInsert(3, 0) = "Address"
aInsert(3, 1) = addstring(xAddress.text)

aInsert(4, 0) = "phone1"
aInsert(4, 1) = addstring(xPhone1.text)

aInsert(5, 0) = "fax"
aInsert(5, 1) = addstring(xFAx.text)

aInsert(6, 0) = "f_balance"
aInsert(6, 1) = addvalue(XF_BALANCE.text)

aInsert(7, 0) = "[Group]"
aInsert(7, 1) = addstring(xGroup.BoundText)

aInsert(8, 0) = "f_date"
aInsert(8, 1) = addDate(xF_date.text)

aInsert(9, 0) = "email"
aInsert(9, 1) = addstring(XEMAIL.text)

aInsert(10, 0) = "disc"
aInsert(10, 1) = addvalue(XDISC.text)

aInsert(11, 0) = "remark"
aInsert(11, 1) = addstring(xRemark.text)

aInsert(12, 0) = "ISONEST"
aInsert(12, 1) = IIf(XISONEST.Value, 1, 0)

aInsert(13, 0) = "IsMonth"
aInsert(13, 1) = IIf(xIsMonth.Value, 1, 0)

aInsert(14, 0) = "[ACC]"
aInsert(14, 1) = addstring(xacc.BoundText)


con.BeginTrans
If xCode.Enabled Then
    xCode.text = RetZero(Newflag("file3_10", "code", con), 4)
    aInsert(0, 1) = addstring(xCode.text)
    cString = CreateInsert(aInsert, "file3_10")
    con.Execute cString
Else
    cString = CreateUpdate(aInsert, "file3_10", " WHERE CODE = " & MyParn(xCode.text))
    con.Execute cString
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
con.RollbackTrans
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
closeCon con
Err.Clear
End Sub

Private Sub xCode_LostFocus()
If xCode.text = "" Then Exit Sub
xCode.text = RetZero(xCode.text, 4)
CardTable.Find "CODE = " & Val(xCode.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Function MYVALID() As Boolean
If xCode.text = "" Then
    MsgBox "«·ﬂÊœ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If xGroup.BoundText = "" Then
    MsgBox "«·„Ã„Ê⁄… ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If xDesca.text = "" Then
    MsgBox "«·≈”„ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

MYVALID = True
End Function



