VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form bankfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·»‰Êﬂ"
   ClientHeight    =   5595
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   8955
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5595
   ScaleWidth      =   8955
   Begin VB.CheckBox xisstop 
      BackColor       =   &H000000FF&
      Caption         =   "Õ”«» „€·ﬁ"
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
      Left            =   135
      TabIndex        =   41
      Top             =   1575
      Width           =   1545
   End
   Begin VB.CheckBox XISSHIP 
      Caption         =   "≈Ìœ«⁄«  ‘Õ‰"
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
      Left            =   135
      TabIndex        =   40
      Top             =   945
      Width           =   1545
   End
   Begin VB.CommandButton cmdGroup 
      Caption         =   "..."
      Height          =   330
      Left            =   3105
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1665
      Width           =   330
   End
   Begin VB.TextBox xPhone 
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
      Height          =   600
      Left            =   180
      MaxLength       =   75
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2655
      Width           =   7080
   End
   Begin VB.TextBox xAddress 
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
      Height          =   600
      Left            =   135
      MaxLength       =   100
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2025
      Width           =   7125
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   8955
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   8955
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   405
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BANK.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   4365
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BANK.frx":014A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdDel 
         Height          =   420
         Left            =   3045
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BANK.frx":058C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   420
         Left            =   1725
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BANK.frx":2E26
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdAdd 
         Height          =   420
         Left            =   5685
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BANK.frx":539F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   7005
         Picture         =   "BANK.frx":794B
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   45
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   8955
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5130
      Width           =   8955
      Begin VB.CommandButton cmdNext 
         Height          =   375
         Left            =   5550
         Picture         =   "BANK.frx":A11E
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "«· «·Ì"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   375
         Left            =   4635
         Picture         =   "BANK.frx":C6DE
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "«·”«»ﬁ"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdLast 
         Height          =   375
         Left            =   3480
         Picture         =   "BANK.frx":ECB1
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "√ŒÌ—"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   375
         Left            =   2610
         Picture         =   "BANK.frx":1138B
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   45
         Width           =   870
      End
   End
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   5940
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   585
      Width           =   1320
   End
   Begin VB.TextBox xDescA 
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
      Left            =   3555
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   945
      Width           =   3705
   End
   Begin VB.TextBox xID 
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
      Left            =   4905
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1305
      Width           =   2355
   End
   Begin MSDataListLib.DataCombo xGroup 
      Height          =   315
      Left            =   3465
      TabIndex        =   3
      Top             =   1665
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1785
      Left            =   180
      TabIndex        =   19
      Top             =   3240
      Width           =   8715
      Begin VB.TextBox xMinusChq 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   180
         MaxLength       =   4
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1305
         Width           =   1410
      End
      Begin VB.TextBox xSUMOUT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   5535
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   585
         Width           =   1545
      End
      Begin VB.TextBox xSUMIN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   135
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   585
         Width           =   1455
      End
      Begin VB.TextBox xf_BAL 
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
         Left            =   5535
         MaxLength       =   12
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   1545
      End
      Begin VB.TextBox xBALANCE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   5535
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   945
         Width           =   1545
      End
      Begin VB.TextBox xPlusChq 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   5535
         MaxLength       =   4
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1305
         Width           =   1545
      End
      Begin VB.TextBox xF_DATE 
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
         Left            =   135
         MaxLength       =   12
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "≈Ã„«·Ï √ . œ :"
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1395
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   7140
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   300
         Width           =   1230
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "≈Ã„«·Ï „”ÕÊ»«  :"
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
         Left            =   7140
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   675
         Width           =   1500
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "≈Ã„«·Ï √ . ﬁ :"
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
         Left            =   7155
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1395
         Width           =   1080
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "—’Ìœ Õ«·Ï  :"
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
         Left            =   7155
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1035
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "≈Ã„«·Ï ≈Ìœ«⁄«  :"
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   675
         Width           =   1305
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ √Ê· «·„œ… :"
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   300
         Width           =   1245
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7335
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   2655
      Width           =   705
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7335
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2115
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7335
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   1755
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "—ﬁ„ «·Õ”«» :"
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
      Left            =   7335
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1395
      Width           =   1065
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7335
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1035
      Width           =   570
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "ﬂÊœ «·»‰ﬂ :"
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
      Left            =   7335
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   630
      Width           =   795
   End
End
Attribute VB_Name = "bankfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim formMode As Byte
Public bedit As Boolean
Dim con As New ADODB.Connection
Dim CardTable As New ADODB.Recordset
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
CmdAdd.Enabled = (nMode = LoadMode And bedit)
cmdDel.Enabled = (nMode = LoadMode And bedit)
cmdSave.Enabled = bedit
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdGroup.Enabled = bedit
cmdFirst.Enabled = (nMode = LoadMode)
xCode.Enabled = Not (nMode = LoadMode)
End Sub
Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FILE5_10"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·»‰ﬂ"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„ «·»‰Êﬂ"
Search3.Show 1
End Sub
Sub mydefine()

xCode.text = RetZero(Val(Newflag("FILE5_10", "CODE", con)))
xDesca.text = ""
xID.text = ""
xF_Bal.text = ""
xF_date.text = ""
xphone.text = ""
xAddress.text = ""
xSUMIN.text = ""
XISSHIP.Value = 0
xisstop.Value = 0
xSUMOUT.text = ""
xPlusChq.text = ""
xMinusChq.text = ""
xBalance.text = ""
xGroup.BoundText = ""
Handlecontrols DefineMode
End Sub
Sub myProc()
CardTable.Find "Code = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
myload
Unload Search3
End Sub
Sub myload()
xCode.text = CardTable!CODE
xDesca.text = TurnValue(CardTable!DESCA, Null, "")
xID.text = TurnValue(CardTable!ID, Null, "")
xF_Bal.text = Format(Val(CardTable!F_BAL & ""), "Fixed")
XISSHIP.Value = IIf(CardTable!ISSHIP, 1, 0)
xisstop.Value = IIf(CardTable!ISSTOP, 1, 0)
xF_date.text = TurnValue(Format(CardTable!F_DATE, "dd-mm-yyyy"), Null, "")
xGroup.BoundText = CardTable!Group & ""
xphone.text = CardTable!phone & ""
xAddress.text = CardTable!Address & ""

'xSUMIN.text = CardTable!sumIn & ""
'xSUMOUT.text = CardTable!sumOut & ""
'xBalance.text = Val(CardTable!sumIn & "") - Val(CardTable!sumOut & "")
'xPlusChq.text = CardTable!PlusChq & ""
'xMinusChq.text = CardTable!MinusChq & ""
'xRecordNumber = "”Ã· " & CARDTABLE.AbsolutePosition + 1 & " „‰ " & myRecordCount
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim nTry As Integer, aInsert(9, 1)
aInsert(0, 0) = "Code"
aInsert(0, 1) = addstring(xCode.text)

aInsert(1, 0) = "desca"
aInsert(1, 1) = addstring(xDesca.text)

aInsert(2, 0) = "address"
aInsert(2, 1) = addstring(xAddress.text)

aInsert(3, 0) = "id"
aInsert(3, 1) = addstring(xID.text)

aInsert(4, 0) = "f_bal"
aInsert(4, 1) = Val(xF_Bal.text)

aInsert(5, 0) = "f_date"
aInsert(5, 1) = DateSq(xF_date.text)

aInsert(6, 0) = "[group]"
aInsert(6, 1) = addvalue(xGroup.BoundText)

aInsert(7, 0) = "phone"
aInsert(7, 1) = addstring(xphone.text)

aInsert(8, 0) = "ISSHIP"
aInsert(8, 1) = XISSHIP.Value

aInsert(9, 0) = "isstop"
aInsert(9, 1) = xisstop.Value

con.BeginTrans
If xCode.Enabled Then
    xCode.text = RetZero(Val(Newflag("FILE5_10", "CODE", con)))
    aInsert(0, 1) = addstring(xCode.text)
    con.Execute CreateInsert(aInsert, "FILE5_10")
Else
    con.Execute CreateUpdate(aInsert, "FILE5_10", " WHERE CODE = " & MyParn(xCode.text))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Function MYVALID() As Boolean
If xCode.text = "" Then
    MsgBox " ”ÃÌ· „”·”· "
    Exit Function
End If
If Not IsDate(xF_date.text) Then
    MsgBox " «—ÌŒ «Ê· «·„œ… €Ì— „”Ã· «Ê €Ì— ’ÕÌÕ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub CmdAdd_Click()
mydefine
If xCode.Enabled Then xCode.SetFocus Else xID.SetFocus
End Sub
Private Sub CmdDel_Click()
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
con.Execute "Delete  From FILE5_10  Where Code = " & MyParn(xCode.text)
CardTable.Requery
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.Find "Code < " & MyParn(xCode.text), , adSearchBackward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveFirst
    myload
Else
    mydefine
End If
End If
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub

Private Sub cmdGroup_Click()
ReDim aPublic(5)
nCode = xGroup.BoundText
aPublic(0) = "FILE5_50"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "ﬂÊœ «·„Ã„Ê⁄…"
aPublic(4) = "≈”„ «·„Ã„Ê⁄…"
aPublic(5) = "„Ã„Ê⁄«  «·»‰Êﬂ"
FlagFrm.bedit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show 1
data1.Refresh
xGroup.BoundText = nCode
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
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
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ »Ì«‰«  «·⁄„Ì· »‰Ã«Õ"
CardTable.Requery
If xCode.Enabled Then
    CmdAdd_Click
Else
    CardTable.Find "code = " & xCode.text, , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload
End If
End Sub
Private Sub CmdUndo_Click()
If (CardTable.EOF And CardTable.BOF) Then
    mydefine
Else
    If xCode.Enabled Then
        CardTable.MoveLast
        myload
    Else
        myload
    End If
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Form_Load()
openCon con
Set CardTable = New ADODB.Recordset
cField1 = myiif("[TYPE] = 2", "VALUE1") & " as SumIn"
cField2 = myiif("[TYPE] = 2", "VALUE2") & " as SumOut"
cField3 = myiif("[TYPE] = 3", "VALUE1") & " as PlusChq"
cField4 = myiif("[TYPE] = 5", "VALUE1") & " as MinusChq"

cString = "SELECT isship,isstop, FILE5_10.CODE,FILE5_10.ID,FILE5_10.DESCA,FILE5_10.[GROUP],FILE5_10.ADDRESS,FILE5_10.PHONE,FILE5_10.F_BAL,FILE5_10.F_date  FROM FILE5_10  " & _
          "  ORDER BY FILE5_10.CODE"

'cString = "SELECT isship,isstop, FILE5_10.CODE,FILE5_10.ID,FILE5_10.DESCA,FILE5_10.[GROUP],FILE5_10.ADDRESS,FILE5_10.PHONE,FILE5_10.F_BAL,FILE5_10.F_date," & cField1 & "," & cField2 & "," & cField3 & "," & cField4 & " FROM FILE5_10 LEFT JOIN BANKMOVE ON FILE5_10.CODE = BANKMOVE.BANK " & _
'          " GROUP BY isship,isstop,FILE5_10.CODE,FILE5_10.ID,FILE5_10.DESCA,FILE5_10.[GROUP],FILE5_10.ADDRESS,FILE5_10.PHONE,FILE5_10.F_BAL,FILE5_10.F_date ORDER BY FILE5_10.CODE"

CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

data1.ConnectionString = strCon
data1.RecordSource = "Select * from file5_50  Order by Desca"

Set xGroup.RowSource = data1
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    mydefine
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub
Private Sub xCode_LostFocus()
xCode.text = RetZero(xCode.text, 6)
CardTable.Find "Code = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub


Private Sub xPhone_GotFocus()
xphone.SelStart = 0
xphone.SelLength = Len(xphone.text)
End Sub
Private Sub xAddress_GotFocus()
xAddress.SelStart = 0
xAddress.SelLength = Len(xAddress.text)
End Sub
Private Sub xCode_GotFocus()
xCode.SelStart = 0
xCode.SelLength = Len(xCode.text)
End Sub
Private Sub xDescA_GotFocus()
xDesca.SelStart = 0
xDesca.SelLength = Len(xDesca.text)
End Sub
Private Sub xID_GotFocus()
xID.SelStart = 0
xID.SelLength = Len(xID.text)
End Sub
Private Sub xMinusChq_GotFocus()
xMinusChq.SelStart = 0
xMinusChq.SelLength = Len(xMinusChq.text)
End Sub
Private Sub xSUMOUT_GotFocus()
xSUMOUT.SelStart = 0
xSUMOUT.SelLength = Len(xSUMOUT.text)
End Sub
Private Sub xSUMIN_GotFocus()
xSUMIN.SelStart = 0
xSUMIN.SelLength = Len(xSUMIN.text)
End Sub
Private Sub xf_BAL_GotFocus()
xF_Bal.SelStart = 0
xF_Bal.SelLength = Len(xF_Bal.text)
End Sub
Private Sub xBALANCE_GotFocus()
xBalance.SelStart = 0
xBalance.SelLength = Len(xBalance.text)
End Sub
Private Sub xPlusChq_GotFocus()
xPlusChq.SelStart = 0
xPlusChq.SelLength = Len(xPlusChq.text)
End Sub
Private Sub xf_Date_GotFocus()
xF_date.SelStart = 0
xF_date.SelLength = Len(xF_date.text)
End Sub
