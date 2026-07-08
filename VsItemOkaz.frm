VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form VsItemOkaz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„ «»⁄… «·„ÊœÌ·« "
   ClientHeight    =   11040
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   2040
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   3750
      Begin VB.Label XTSAL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1620
         Width           =   1725
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ﬁÌ„… «·„»Ì⁄« "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1620
         Width           =   1320
      End
      Begin VB.Label xRate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   180
         Width           =   555
      End
      Begin VB.Label xTBal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   900
         Width           =   1725
      End
      Begin VB.Label xTSalCOST 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1260
         Width           =   1725
      End
      Begin VB.Label xTPurch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   540
         Width           =   1725
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "‰”»… «·„»Ì⁄« "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ﬂ·›… „»Ì⁄« "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1260
         Width           =   1320
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ﬁÌ„… «·—’Ìœ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   900
         Width           =   1320
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ﬁÌ„… „‘ —Ì« "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   540
         Width           =   1320
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "⁄œœ «·„ÊœÌ·« "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   180
         Width           =   1320
      End
      Begin VB.Label xCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   180
         Width           =   645
      End
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2055
      Width           =   1155
   End
   Begin VB.CommandButton Cmd_Print 
      Caption         =   "ÿ»«⁄…"
      Height          =   420
      Left            =   1350
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2055
      Width           =   1155
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "⁄—÷"
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
      Left            =   2610
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2055
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   2250
      Left            =   3810
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11235
      Begin VB.CommandButton CMD_CHANG 
         Caption         =   " €Ì— ”⁄— «·«Êﬂ«“ÌÊ‰"
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
         Left            =   1395
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1305
         Width           =   1740
      End
      Begin VB.TextBox XOKAZ 
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
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1320
         Width           =   1545
      End
      Begin VB.TextBox xBarCode 
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
         Left            =   7890
         MaxLength       =   15
         TabIndex        =   22
         Top             =   1665
         Width           =   1545
      End
      Begin VB.TextBox xModelFact 
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
         Left            =   3195
         MaxLength       =   15
         TabIndex        =   15
         Top             =   930
         Width           =   1545
      End
      Begin VB.TextBox xDescItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EAEAEA&
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
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   930
         Width           =   3015
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
         Left            =   7620
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   6000
         TabIndex        =   8
         Top             =   1308
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   6000
         TabIndex        =   9
         Top             =   952
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   1305
         TabIndex        =   10
         Top             =   180
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   315
         Left            =   6000
         TabIndex        =   18
         Top             =   596
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   1305
         TabIndex        =   20
         Top             =   555
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… «Êﬂ«“ÌÊ‰ "
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
         Index           =   4
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1350
         Width           =   1170
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "»«—ﬂÊœ :"
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
         Height          =   195
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1725
         Width           =   585
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬁ”„ :"
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   615
         Width           =   1230
      End
      Begin VB.Label Label5 
         Caption         =   "«·„Ê—œ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   585
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„’‰⁄"
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
         Index           =   2
         Left            =   1755
         TabIndex        =   17
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ „ÊœÌ· :"
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
         Height          =   195
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   990
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "„Ã„Ê⁄… „ÊœÌ·«  :"
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
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1356
         Width           =   1410
      End
      Begin VB.Label Label3 
         Caption         =   "«·„’‰⁄ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   974
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ê”„ :"
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»⁄œ  «—ÌŒ :"
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
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   270
         Width           =   765
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10710
      Width           =   15120
      _ExtentX        =   26670
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   1485
      Top             =   45
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "VsItemOkaz.frx":0000
      Height          =   7410
      Left            =   45
      TabIndex        =   7
      Top             =   2475
      Width           =   15000
      _cx             =   26458
      _cy             =   13070
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      BackColorSel    =   16777088
      ForeColorSel    =   4194368
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   2
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   1188
      Top             =   45
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
      Left            =   891
      Top             =   45
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   297
      Top             =   45
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
   Begin MSAdodcLib.Adodc data5 
      Height          =   330
      Left            =   0
      Top             =   45
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
   Begin MSAdodcLib.Adodc data6 
      Height          =   330
      Left            =   594
      Top             =   45
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
   Begin VSFlex7Ctl.VSFlexGrid VsStore 
      Bindings        =   "VsItemOkaz.frx":0014
      Height          =   660
      Left            =   135
      TabIndex        =   25
      Top             =   9945
      Visible         =   0   'False
      Width           =   14865
      _cx             =   26220
      _cy             =   1164
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777088
      ForeColorSel    =   4194368
      BackColorBkg    =   16777215
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   2
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
   Begin VB.Label xmodelno 
      Alignment       =   1  'Right Justify
      Caption         =   "Label7"
      Height          =   240
      Left            =   4095
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2205
      Visible         =   0   'False
      Width           =   555
   End
End
Attribute VB_Name = "VsItemOkaz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cMyModel As String
Dim con As New ADODB.Connection
Dim StoreTable As New ADODB.Recordset
Private Sub CMD_CHANG_Click()
If XOKAZ.Text = "" Or Not IsNumeric(XOKAZ.Text) Then Exit Sub
If MsgBox(" €Ì— ‰”»… «·«Êﬂ«“ÌÊ‰", vbYesNo + vbDefaultButton2) = vbYes Then
With grid1
    For I = 3 To .Rows - 1
        DoFixPrice2 .TextMatrix(I, 0), Val(XOKAZ.Text), I
    Next I
End With
End If
End Sub

Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï  ⁄«„·«  «·„ÊœÌ·«  "
    cHead2 = " Õ Ï   «—ÌŒ " & Format(xDate1.Text, "DD-MM-YYYY")
    Load PrintGrd
    PrintGrd.doprint Me.grid1, 0.75, -2, cHead1, cHead2, , False, True, 6
    PrintGrd.Show 1
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
    openCon con
    StoreTable.Open "FILE0_40", con, adOpenStatic, adLockReadOnly, adCmdTable
    xDate1.Text = Format(Date, "dd-mm-yyyy")
    
    data1.ConnectionString = strCon
    data1.RecordSource = "Select Code,DescA From File1_10SC order by Desca"
    Set xSection.RowSource = data1
    xSection.ListField = "Desca"
    xSection.BoundColumn = "Code"
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "Select Code,DescA From File4_10 order by Desca"
    Set xSupp.RowSource = DATA2
    xSupp.ListField = "Desca"
    xSupp.BoundColumn = "Code"
    
    data3.ConnectionString = strCon
    data3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
    Set xGroup.RowSource = data3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    data4.ConnectionString = strCon
    data4.RecordSource = "Select mosm ,descA From mosm ORDER BY date "
    Set xMosm.RowSource = data4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm
    data5.ConnectionString = strCon
    data5.RecordSource = "Select code ,desca From fact ORDER BY desca "
    Set xFact.RowSource = data5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    
    Set grid1.DataSource = data6
    data6.ConnectionString = strCon
    
    grid1.Rows = 2
    grid1.Cols = 16
    
    
    FixGrid
    FixGridStore
End Sub
Private Sub myload()
    Dim cwhere As String

    cField5 = myiif("type = 'z' ", "([out] - [in]) ") & " AS t_comp , "
    cField6 = myiif("type = '2' ", "([IN] - [OUT]) ") & " AS t_purch, "
    cField7 = myiif("type = '7' ", "([out] - [in]) ") & " AS t_ret  , "
    cField8 = myiif("type = '6' ", "([out] - [in]) ") & " AS t_sal  , "
    cField14 = myiif2("type = '2' ", "[date]", "MIN") & " AS f_date , "
    cField16 = myiif2("type = '2' ", "[date]  ", "MAX") & " AS l_date , "
    '                   0                   1               2           3                   4
    cString = " SELECT file1_10.ModelNo, FACT.desca , file1_10.mosm, file1_10.MODELFACT0, file1_10.desca ,   " & _
            cField5 & cField6 & cField7 & cField8 & _
            " sum([in] - out) as bal ,  " & _
            " ' ' as rate , min(file1_10.price) , min(file1_10.cost) , sum((file1_11.[in] - file1_11.out) * file1_10.cost ) as costbal , min(file1_10.price3) , MIN(OKAZ   ) " & _
            " FROM    FILE1_10 INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM WHERE FILE1_10.MODELNO IS NOT NULL "
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xModelFact.Text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELFACT0 = " & MyParn(xModelFact.Text)
    If xBarCode.Text <> "" Then cStr2 = cStr2 & " AND FILE1_10.item = " & xBarCode.Text
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xSection.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[Section] = " & xSection.BoundText
    If IsDate(xDate1.Text) Then cStr2 = cStr2 & " AND FILE1_11.[DATE] <= " & DateSq(xDate1.Text)
    cString = cString & cStr2 & " GROUP BY file1_10.[SECTION],file1_10.ModelNo, FACT.desca , file1_10.mosm, file1_10.MODELFACT0, file1_10.desca  "
    data6.RecordSource = cString
    data6.Refresh
    FixGrid

'   CalcTot
End Sub
Sub FixGrid()
    Dim CodeModelTable As New ADODB.Recordset
    CodeModelTable.Open "ModelSupp", con, adOpenKeyset, adLockReadOnly, adCmdTable
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 2
    .RowHidden(1) = True
    .RowHeight(0) = 1000
    .WordWrap = True
    .ColHidden(0) = True
    .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
    .RowHidden(1) = True
    .FrozenCols = 4
    .TextMatrix(0, 0) = "„ÊœÌ·"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .TextMatrix(0, 2) = "„Ê”„"
    .TextMatrix(0, 3) = "„ÊœÌ· "
    .TextMatrix(0, 4) = "«·’‰›"
    
    .TextMatrix(0, 5) = "»«—ﬂÊœ"
    .TextMatrix(0, 6) = "„‘ —Ì« "
    .TextMatrix(0, 7) = "„— Ã⁄« "
    .TextMatrix(0, 8) = "„»Ì⁄« "
    .TextMatrix(0, 9) = "—’Ìœ"
    
    .TextMatrix(0, 10) = "‰”»… »Ì⁄"
    .TextMatrix(0, 11) = " ”⁄— »Ì⁄"
    .TextMatrix(0, 12) = "”⁄— „’‰⁄"
    .TextMatrix(0, 13) = " ﬁÌ„… —’Ìœ"
    .TextMatrix(0, 14) = " ”⁄— √Êﬂ«“ÌÊ‰"
    .TextMatrix(0, 15) = " ‰”»… ·√Êﬂ«“ÌÊ‰"
    
    
    .ColWidth(0) = 0
    .ColWidth(1) = 1200
    .ColWidth(2) = 500
    .ColWidth(3) = 1000
    .ColWidth(4) = 3000
    
    .ColWidth(5) = 1000
    .ColWidth(6) = 850
    .ColWidth(7) = 870
    .ColWidth(8) = 800
    .ColWidth(9) = 800
    
    .ColWidth(10) = 700
    .ColWidth(11) = 700
    .ColWidth(12) = 700
    .ColWidth(13) = 0
    
    .ColWidth(14) = 1000
    .ColWidth(15) = 1000
    .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = 4
    
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTString
    .ColDataType(4) = flexDTString

    For nCol = 5 To .Cols - 1
        .ColDataType(nCol) = flexDTDouble
    Next nCol

    .Editable = flexEDNone
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
    .Subtotal flexSTClear
    
    If .Rows > 2 Then
        For I = 2 To .Rows - 1
            If (Val(.TextMatrix(I, 6)) = Val(.TextMatrix(I, 7))) And Val(.TextMatrix(I, 7)) > 0 Then
                .Cell(flexcpBackColor, I, 0, I, .Cols - 1) = &HC0C0FF
            End If
            If (Val(.TextMatrix(I, 6)) + Val(.TextMatrix(I, 7))) > 0 Then nRat = Val(.TextMatrix(I, 8)) / (Val(.TextMatrix(I, 7)) + Val(.TextMatrix(I, 6))) * 100
            .TextMatrix(I, 10) = Format(nRat, "##0.00")
            .TextMatrix(I, 13) = Format(Val(.TextMatrix(I, 13)), "##0.00")
        Next I
    End If
    .Subtotal flexSTSum, -1, 5, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 6, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 7, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 8, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 9, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 13, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    If .Rows > 2 Then
        xCount.Caption = .Rows - 3
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
    Else
        xCount.Caption = ""
    End If
    
    
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Row > 2 Then
        If Col = 14 Then
            con.Execute " update file1_10 set price3 = " & Val(grid1.TextMatrix(Row, 14)) & " , isdate = getdate() where modelno = " & MyParn(grid1.TextMatrix(Row, 0))
        End If
        If Col = 15 Then
            DoFixPrice2 grid1.TextMatrix(Row, 0), Val(grid1.TextMatrix(Row, 15)), Row
        End If
    End If
End Sub
Sub DoFixPrice2(cModel, nOkaz, nRow)
    Dim SubItemTable As New ADODB.Recordset
    Dim nRem As Double, nPrice As Double
    Dim nFPrice As Double
    
    cStr1 = " SELECT * FROM FILE1_10  WHERE MODELNO = " & MyParn(cModel)
    SubItemTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not SubItemTable.EOF
            nPrice = Val(SubItemTable!price & "") * (100 - nOkaz) / 100
            nRem = nPrice - Int(nPrice)
            Select Case nRem
                Case Is = 0
                    nPrice = Int(nPrice)
                Case Is <= 0.5
                    nPrice = Int(nPrice) + 0.5
                Case Is < 1
                    nPrice = Int(nPrice) + 1
            End Select
            If nPrice > SubItemTable!price Then nPrice = SubItemTable!price
            If nFPrice = 0 Then nFPrice = nPrice
            con.Execute " UPDATE FILE1_10 SET PRICE3 = " & Val(nPrice) & " , isdate = getdate() WHERE ITEM = " & MyParn(SubItemTable!Item)
            SubItemTable.MoveNext
        Loop
        con.Execute " UPDATE FILE1_10 SET OKAZ          = " & Val(nOkaz) & " , isdate = getdate() WHERE MODELNO = " & MyParn(cModel)
        grid1.TextMatrix(nRow, 14) = Format(nFPrice, "#0.00")
        grid1.TextMatrix(nRow, 15) = Format(nOkaz, "#0")
'       MsgBox " „ Ÿ»ÿ ”⁄— «·«Êﬂ«“ÌÊ‰ ··„ÊœÌ·"

End Sub

Private Sub grid1_DBLClick()
    If grid1.Col = 4 Then
        lExirSave = True
        If Val(GetDesca("SELECT COUNT (CODE) FROM ModelSupp WHERE MODELNO = " & MyParn(xmodelno.Caption)) & "") > 1 Then
            FrmListSupp.Show 1
        Else
            cMyModel = GetDesca("SELECT MODEL FROM ModelSupp WHERE MODELNO = " & MyParn(xmodelno.Caption))
        End If
        Load ITEMS
        ITEMS.xModel.Text = cMyModel
        ITEMS.CardTable.Find "MODEL = " & MyParn(ITEMS.xModel.Text), , adSearchForward, adBookmarkFirst
        If Not ITEMS.CardTable.EOF Then
            ITEMS.myload
            ITEMS.Show 1
        End If
    End If
    
    If grid1.Col > 6 And grid1.Col < 10 Then
        aRet = aGetDesca(" select modelno , modelfact , fact , mosm , desca from file1_10h where modelno = " & MyParn(grid1.TextMatrix(grid1.Row, 0)))
        If UBound(aRet) > 0 Then
            VsModelNoLook.xmodelno.Text = aRet(1)
            VsModelNoLook.xFact.BoundText = aRet(3)
            VsModelNoLook.xModelFact.Text = RetZero(aRet(2))
            VsModelNoLook.xMosm.BoundText = aRet(4)
            VsModelNoLook.xDesca.Text = aRet(5)
            VsModelNoLook.myloadGrdModel
            VsModelNoLook.DataGrdModel
            VsModelNoLook.Show
        End If
    End If
End Sub
Private Sub grid1_EnterCell()
    xmodelno.Caption = grid1.TextMatrix(grid1.Row, 0)
    With grid1
        If .Col = 15 Or .Col = 14 Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
End Sub
Sub ShowBalStore(cModelNo)
    Dim datatable As New ADODB.Recordset
    cField1 = myiif("type = '2' OR type = '7' ", "([IN] - [OUT]) ") & " AS T_Purch , "
    cField2 = myiif("type = 'F' OR type = 'T' ", "([IN]-[out]) ") & " AS T_Trans , "
    cField3 = myiif("type = '6' ", "([out]) ") & " AS T_Sal , "
    cField4 = myiif("type = 'z' ", "([out] - [in]) ") & " AS T_Comp  "
    cString = " SELECT FILE1_11.STORE , sum([in] - out) as bal , " & _
            cField1 & cField2 & cField3 & cField4 & _
            " FROM (file1_10 INNER JOIN file1_11 ON file1_10.ITEM = file1_11.item) where modelno = " & MyParn(cModelNo)
    cString = cString & " GROUP BY FILE1_11.STORE "
    datatable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    With VsStore
        For nCol = 1 To .Cols - 1
            datatable.Find " STORE = " & MyParn(.TextMatrix(0, nCol)), , adSearchForward, adBookmarkFirst
            If Not datatable.EOF Then
                .TextMatrix(2, nCol) = Format(datatable!BAL, "#0")
                .TextMatrix(3, nCol) = Format(datatable!T_PURCH, "#0")
                .TextMatrix(4, nCol) = Format(datatable!T_TRANS, "#0")
                .TextMatrix(5, nCol) = Format(datatable!t_sal, "#0")
                .TextMatrix(6, nCol) = Format(datatable!T_COMP, "#0")
            Else
                .TextMatrix(2, nCol) = ""
                .TextMatrix(3, nCol) = ""
                .TextMatrix(4, nCol) = ""
                .TextMatrix(5, nCol) = ""
                .TextMatrix(6, nCol) = ""
            End If
        Next nCol
    End With
End Sub
Sub FixGridStore()
    With VsStore
        .Rows = 7
        .FixedRows = 2
        .RowHidden(0) = True
        .Cols = 1
        .TextMatrix(1, 0) = "«·›—⁄"
        .TextMatrix(2, 0) = "«·—’Ìœ"
        .TextMatrix(3, 0) = "„‘ —Ì« "
        .TextMatrix(4, 0) = " ÕÊÌ·« "
        .TextMatrix(5, 0) = "„»Ì⁄« "
        .TextMatrix(6, 0) = " ”ÊÌ… Ã—œ"
        .ColWidth(0) = 2000
        StoreTable.MoveFirst
        Do While Not StoreTable.EOF
            .Cols = .Cols + 1
            .TextMatrix(0, .Cols - 1) = StoreTable!code
            .TextMatrix(1, .Cols - 1) = StoreTable!DESCA
            .ColWidth(.Cols - 1) = 1200
            StoreTable.MoveNext
        Loop
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    End With
End Sub
Sub CalcTot()
    Dim cF1 As String
    cString = " SELECT "
    cF1 = myiif("(type = '2' OR type = '7') ", "([IN] - [OUT]) * FILE1_11.PRICE * ((100 - FILE1_11.DISCOUNT ) /100)") & " AS t_purch , "
    cF1 = cF1 & myiif("type = '6' ", "([OUT]) * FILE1_11.PRICE * ((100 - FILE1_11.DISCOUNT ) /100)") & " AS t_PRICE , "
    cF1 = cF1 & myiif("type = '6' ", "([out] * FILE1_10.COST) ") & " AS t_SALCOST , "
    '                   0                   1                            2               3                   4                              5
    cString = cString & cF1 & _
            " sum(([in] - out) * file1_10.cost ) as Tbal   " & _
            " FROM (file1_10 INNER JOIN file1_11 ON file1_10.ITEM = file1_11.item) where modelno is not null "
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xModelFact.Text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELFACT0 = " & MyParn(xModelFact.Text)
    If xBarCode.Text <> "" Then cStr2 = cStr2 & " AND FILE1_10.item = " & xBarCode.Text
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xSection.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[Section] = " & xSection.BoundText

    If IsDate(xDate1.Text) Then cStr2 = cStr2 & " AND FILE1_11.[DATE] <= " & DateSq(xDate1.Text)
    cString = cString & cStr2
    aRet = aGetDesca(cString)
    If UBound(aRet) > 0 Then
        xTPurch.Caption = Format(aRet(1), "#0.00")
        xTSalCOST.Caption = Format(aRet(3), "#0.00")
        XTSAL.Caption = Format(aRet(2), "#0.00")
        xTBal.Caption = Format(aRet(4), "#0.00")
        If Val(xTPurch.Caption) <> 0 Then xRate.Caption = Format(Val(xTSalCOST.Caption) / Val(xTPurch.Caption) * 100, "#0.00")
    End If
End Sub
