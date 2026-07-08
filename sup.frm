VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form supfrm 
   Caption         =   "»Ì«‰«  «·„Ê—œÌ‰"
   ClientHeight    =   5610
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   9225
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
   ScaleHeight     =   5610
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExit 
      Height          =   420
      Left            =   1980
      MaskColor       =   &H00FFFFFF&
      Picture         =   "sup.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   90
      UseMaskColor    =   -1  'True
      Width           =   1140
   End
   Begin VB.CommandButton CmdSave 
      Height          =   420
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      Picture         =   "sup.frx":014A
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Õ›Ÿ"
      Top             =   90
      UseMaskColor    =   -1  'True
      Width           =   1140
   End
   Begin VB.CommandButton CmdDel 
      Height          =   420
      Left            =   4260
      MaskColor       =   &H00FFFFFF&
      Picture         =   "sup.frx":058C
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "Õ–›"
      Top             =   90
      UseMaskColor    =   -1  'True
      Width           =   1140
   End
   Begin VB.CommandButton CmdUndo 
      Height          =   420
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "sup.frx":2E26
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   38
      TabStop         =   0   'False
      ToolTipText     =   " —«Ã⁄"
      Top             =   90
      UseMaskColor    =   -1  'True
      Width           =   1140
   End
   Begin VB.CommandButton CmdAdd 
      Height          =   420
      Left            =   6540
      MaskColor       =   &H00FFFFFF&
      Picture         =   "sup.frx":539F
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "«÷«›…"
      Top             =   90
      UseMaskColor    =   -1  'True
      Width           =   1140
   End
   Begin VB.CommandButton CmdInform 
      Height          =   420
      Left            =   7680
      Picture         =   "sup.frx":794B
      Style           =   1  'Graphical
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "«” ⁄·«„"
      Top             =   90
      Width           =   1140
   End
   Begin VB.CommandButton cmdGroup 
      Caption         =   "..."
      Height          =   330
      Left            =   3780
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1755
      Width           =   375
   End
   Begin VB.TextBox xDisc2 
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
      Left            =   315
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3735
      Width           =   1500
   End
   Begin VB.TextBox xDisc1 
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
      Left            =   5850
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3735
      Width           =   1545
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   360
      Top             =   2655
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
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
      Caption         =   "—’Ìœ «›  «ÕÌ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   4095
      Width           =   8925
      Begin VB.TextBox xF_bal1 
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
         Left            =   5670
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   180
         Width           =   1545
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
         Left            =   5670
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Tag             =   "date"
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox xF_bal2 
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
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "„œÌ‰ :"
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
         Left            =   7425
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   270
         Width           =   480
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
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   615
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "œ«∆‰ :"
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
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   225
         Width           =   435
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
      ScaleWidth      =   9225
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5145
      Width           =   9225
      Begin VB.CommandButton cmdFirst 
         Height          =   375
         Left            =   2700
         Picture         =   "sup.frx":A11E
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdLast 
         Height          =   375
         Left            =   3570
         Picture         =   "sup.frx":C7FD
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "√ŒÌ—"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdPerviious 
         Height          =   375
         Left            =   4725
         Picture         =   "sup.frx":EED7
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "«·”«»ﬁ"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdNext 
         Height          =   375
         Left            =   5640
         Picture         =   "sup.frx":114AA
         Style           =   1  'Graphical
         TabIndex        =   32
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
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   9225
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   9225
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
      Height          =   510
      Left            =   315
      MaxLength       =   200
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2115
      Width           =   7080
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
      Left            =   2700
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1395
      Width           =   4695
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
      Height          =   330
      Left            =   5850
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   3375
      Width           =   1545
   End
   Begin VB.TextBox xPhone2 
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
      Left            =   4275
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3015
      Width           =   1545
   End
   Begin VB.TextBox xFileNo 
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
      Left            =   2700
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2655
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
      Height          =   330
      Left            =   6075
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   675
      Width           =   1320
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
      Left            =   2700
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1035
      Width           =   4695
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
      Left            =   5850
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3015
      Width           =   1545
   End
   Begin VB.TextBox xPhone3 
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
      Left            =   2700
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   3015
      Width           =   1545
   End
   Begin MSDataListLib.DataCombo xgroup 
      Height          =   315
      Left            =   4185
      TabIndex        =   3
      Top             =   1755
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Œ’„ œ«Œ·Ì :"
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
      Left            =   1890
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   3870
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Œ’„ Œ«—ÃÌ :"
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
      Left            =   7530
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   3840
      Width           =   1110
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "›«ﬂ” :"
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
      Left            =   7515
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3510
      Width           =   570
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
      Left            =   7515
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1800
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
      Left            =   7515
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2115
      Width           =   645
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "«·„œÌ— «·„”∆Ê· :"
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
      Left            =   7515
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1485
      Width           =   1290
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·„·› «·÷—Ì»Ì :"
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
      Left            =   7515
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2700
      Width           =   1620
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
      Left            =   7515
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1110
      Width           =   570
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
      Left            =   7515
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   735
      Width           =   360
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
      Left            =   7515
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   3060
      Width           =   705
   End
End
Attribute VB_Name = "supfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean
Public myFlag As Integer
Dim con As New ADODB.Connection
Dim formMode As Byte, cTableName As String, cGroupname As String
Dim CardTable As New ADODB.Recordset
Const LoadMode = 1, DefineMode = 2
Private Sub cmdGroup_Click()
Dim myPublic(6)
myPublic(0) = "FILE4_50"
myPublic(1) = "Code"
myPublic(2) = "Desca"
myPublic(3) = "ﬂÊœ «·„Ã„Ê⁄…"
myPublic(4) = "≈”„ «·„Ã„Ê⁄…"
myPublic(5) = "„Ã„Ê⁄«  «·„Ê—œÌ‰"
FlagFrm2.bEdit = True
FlagFrm2.myPublic = myPublic
FlagFrm2.Show 1
data1.Refresh
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
openCon con
data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FILE4_50"
Set xgroup.RowSource = data1
xgroup.ListField = "Desca"
xgroup.BoundColumn = "Code"

CardTable.Open "Select * From FILE4_10 order by code", con, adOpenKeyset, adLockReadOnly, adCmdText
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    myDefine
End If
End Sub
Private Sub CmdAdd_Click()
myDefine
XCODE.SetFocus
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    con.BeginTrans
    con.Execute "Delete  From FILE4_10  Where code = " & MyParn(XCODE.Text)
    con.CommitTrans
    CardTable.Requery
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "code < " & MyParn(XCODE.Text), , adSearchBackward, adBookmarkLast
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
    con.RollbackTrans
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ »Ì«‰«  «·⁄„Ì· »‰Ã«Õ"
CardTable.Requery
If XCODE.Enabled Then
    CmdAdd_Click
Else
    CardTable.Find "code = " & XCODE.Text, , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload
End If

End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.EOF And CardTable.BOF Then
    myDefine
Else
    If XCODE.Enabled Then
        CardTable.MoveLast
    Else
        CardTable.Find "code = " & MyParn(XCODE.Text), , adSearchForward, adBookmarkFirst
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
Sub Handlecontrols(nMode)
cmdAdd.Enabled = (nMode = LoadMode And bEdit)
CmdDel.Enabled = (nMode = LoadMode And bEdit)
CmdSave.Enabled = bEdit
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
XCODE.Enabled = Not (nMode = LoadMode)
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FILE4_10"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·«”„"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "≈” ⁄·«„ «·⁄„·«¡"
Search3.Show 1
End Sub
Sub myDefine()
XCODE.Text = RetZero(Newflag("FILE4_10", "code"), 6)
xDesca.Text = ""
xManager.Text = ""
xFileNo.Text = ""
xgroup.BoundText = ""
xAddress.Text = ""
xPhone1.Text = ""
xPhone2.Text = ""
xPhone3.Text = ""
xFAx.Text = ""
xF_Date.Text = ""
xF_bal1.Text = ""
xF_bal2.Text = ""
xDisc1.Text = ""
xDisc2.Text = ""
Handlecontrols DefineMode
End Sub
Sub myload()
XCODE.Text = CardTable!CODE & ""
xDesca.Text = CardTable!desca
xManager.Text = CardTable!Manager & ""
xAddress.Text = CardTable!ADDRESS & ""
xFileNo.Text = CardTable!fileNo & ""
xPhone1.Text = CardTable!PHONE1 & ""
xPhone2.Text = CardTable!phone2 & ""
xPhone3.Text = CardTable!phone3 & ""
xFAx.Text = CardTable!Fax & ""
xF_Date.Text = Format(CardTable!F_dATE, "dd-mm-yyyy")
If CardTable!F_Balance > 0 Then xF_bal1.Text = Format(CardTable!F_Balance, "##0.00")
If CardTable!F_Balance < 0 Then xF_bal2.Text = Format(Abs(CardTable!F_Balance), "##0.00")
xDisc1.Text = Myvalue(CardTable!DISC1, "##0.00")
xDisc2.Text = Myvalue(CardTable!DISC2, "##0.00")
xgroup.BoundText = CardTable!Group & ""
xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim aInsert(11, 1)
aInsert(0, 0) = "Code"
aInsert(0, 1) = addstring(XCODE.Text)

aInsert(1, 0) = "[Desca]"
aInsert(1, 1) = addstring(xDesca.Text)

aInsert(2, 0) = "Manager"
aInsert(2, 1) = addstring(xManager.Text)

aInsert(3, 0) = "Phone1"
aInsert(3, 1) = addstring(xPhone1.Text)

aInsert(4, 0) = "Phone2"
aInsert(4, 1) = addstring(xPhone2.Text)

aInsert(5, 0) = "Phone3"
aInsert(5, 1) = addstring(xPhone3.Text)

aInsert(6, 0) = "Fax"
aInsert(6, 1) = addstring(xFAx.Text)

aInsert(7, 0) = "F_date"
aInsert(7, 1) = addDate(xF_Date.Text)

If Val(xF_bal1.Text) > 0 Then
    aInsert(8, 0) = "F_Balance"
    aInsert(8, 1) = Val(xF_bal1.Text)
Else
    aInsert(8, 0) = "F_Balance"
    aInsert(8, 1) = Val(xF_bal2.Text) * -1
End If

aInsert(9, 0) = "[Group]"
aInsert(9, 1) = addstring(xgroup.BoundText)

aInsert(10, 0) = "fileNo"
aInsert(10, 1) = addstring(xFileNo.Text)

aInsert(11, 0) = "Address"
aInsert(11, 1) = addstring(xAddress.Text)

On Error GoTo myerror
con.BeginTrans
If XCODE.Enabled Then
    XCODE.Text = RetZero(Val(Newflag("FILE4_10", "code")))
    aInsert(0, 1) = addstring(XCODE.Text)
    con.Execute CreateInsert(aInsert, "FILE4_10")
Else
    con.Execute CreateUpdate(aInsert, "FILE4_10", " where code = " & addstring(XCODE.Text))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
   CardTable.Find "CODE = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
   myload
   Search3.Hide
End Sub
Private Sub Form_Unload(Cancel As Integer)
CardTable.Close
Set CardTable = Nothing
closeCon con
On Error Resume Next
Unload Search3
Set Search3 = Nothing
Err.Clear
End Sub
Private Sub xCode_LostFocus()
XCODE.BackColor = &H80000005
If XCODE.Text = "" Then Exit Sub
XCODE.Text = RetZero(XCODE.Text, 6)
CardTable.Find "CODE = " & MyParn(XCODE.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Function MYVALID() As Boolean
If XCODE.Text = "" Then
    MsgBox "«·ﬂÊœ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If xgroup.BoundText = "" Then
    MsgBox " ”ÃÌ· «·„Ã„Ê⁄… "
    Exit Function
End If

If xDesca.Text = "" Then
    MsgBox "«·≈”„ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If Not IsDate(xF_Date.Text) Then
    MsgBox " «—ÌŒ «Ê· «·„œ… ÷—Ê—Ì"
    Exit Function
End If
MYVALID = True
End Function
Private Sub xDisc2_LostFocus()
xDisc2.BackColor = &H80000005
End Sub
Private Sub xDisc1_LostFocus()
xDisc1.BackColor = &H80000005
End Sub
Private Sub xBalancePlus_LostFocus()
xBalancePlus.BackColor = &H80000005
End Sub
Private Sub xf_Date_LostFocus()
xF_Date.BackColor = &H80000005
End Sub
Private Sub xBalanceMinus_LostFocus()
xBalanceMinus.BackColor = &H80000005
End Sub
Private Sub xAddress_LostFocus()
xAddress.BackColor = &H80000005
End Sub
Private Sub xManager_LostFocus()
xManager.BackColor = &H80000005
End Sub
Private Sub xFAx_LostFocus()
xFAx.BackColor = &H80000005
End Sub
Private Sub xPhone2_LostFocus()
xPhone2.BackColor = &H80000005
End Sub
Private Sub xfileNo_LostFocus()
xFileNo.BackColor = &H80000005
End Sub
Private Sub xDescA_LostFocus()
xDesca.BackColor = &H80000005
End Sub
Private Sub xPhone1_LostFocus()
xPhone1.BackColor = &H80000005
End Sub
Private Sub xPhone3_LostFocus()
xPhone3.BackColor = &H80000005
End Sub
Private Sub xGroup_LostFocus()
xgroup.BackColor = &H80000005
End Sub
Private Sub xDisc2_GotFocus()
xDisc2.SelStart = 0
xDisc2.SelLength = Len(xDisc2.Text)
xDisc2.BackColor = &HC0FFFF
End Sub
Private Sub xDisc1_GotFocus()
xDisc1.SelStart = 0
xDisc1.SelLength = Len(xDisc1.Text)
xDisc1.BackColor = &HC0FFFF
End Sub
Private Sub xBalancePlus_GotFocus()
xBalancePlus.SelStart = 0
xBalancePlus.SelLength = Len(xBalancePlus.Text)
xBalancePlus.BackColor = &HC0FFFF
End Sub
Private Sub xf_Date_GotFocus()
xF_Date.SelStart = 0
xF_Date.SelLength = Len(xF_Date.Text)
xF_Date.BackColor = &HC0FFFF
End Sub
Private Sub xBalanceMinus_GotFocus()
xBalanceMinus.SelStart = 0
xBalanceMinus.SelLength = Len(xBalanceMinus.Text)
xBalanceMinus.BackColor = &HC0FFFF
End Sub
Private Sub xAddress_GotFocus()
xAddress.SelStart = 0
xAddress.SelLength = Len(xAddress.Text)
xAddress.BackColor = &HC0FFFF
End Sub
Private Sub xManager_GotFocus()
xManager.SelStart = 0
xManager.SelLength = Len(xManager.Text)
xManager.BackColor = &HC0FFFF
End Sub
Private Sub xFAx_GotFocus()
xFAx.SelStart = 0
xFAx.SelLength = Len(xFAx.Text)
xFAx.BackColor = &HC0FFFF
End Sub
Private Sub xPhone2_GotFocus()
xPhone2.SelStart = 0
xPhone2.SelLength = Len(xPhone2.Text)
xPhone2.BackColor = &HC0FFFF
End Sub
Private Sub xfileNo_GotFocus()
xFileNo.SelStart = 0
xFileNo.SelLength = Len(xFileNo.Text)
xFileNo.BackColor = &HC0FFFF
End Sub
Private Sub xCode_GotFocus()
XCODE.SelStart = 0
XCODE.SelLength = Len(XCODE.Text)
XCODE.BackColor = &HC0FFFF
End Sub
Private Sub xDescA_GotFocus()
xDesca.SelStart = 0
xDesca.SelLength = Len(xDesca.Text)
xDesca.BackColor = &HC0FFFF
End Sub
Private Sub xPhone1_GotFocus()
xPhone1.SelStart = 0
xPhone1.SelLength = Len(xPhone1.Text)
xPhone1.BackColor = &HC0FFFF
End Sub
Private Sub xPhone3_GotFocus()
xPhone3.SelStart = 0
xPhone3.SelLength = Len(xPhone3.Text)
xPhone3.BackColor = &HC0FFFF
End Sub
Private Sub xgroup_GotFocus()
xgroup.BackColor = &HC0FFFF
End Sub
Private Sub xf_Date_Validate(Cancel As Boolean)
With xF_Date
If (Not IsDate(.Text)) And Trim(.Text) <> "" Then .Text = ""
.Text = Format(.Text, "dd-mm-yyyy")
End With
End Sub
Private Sub xGroup_Validate(Cancel As Boolean)
If Not xgroup.MatchedWithList Then xgroup.BoundText = ""
End Sub
