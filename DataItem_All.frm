VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form DataItem_All 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ›’Ì·Ï »Ì«‰«  «·„ÊœÌ·« "
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
   Begin VB.Frame Frame5 
      Height          =   555
      Left            =   3780
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   315
      Width           =   2895
      Begin VB.CheckBox chkBarCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰ »«—ﬂÊœ œ«Œ·Ì"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   675
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   180
         Width           =   2040
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1230
      Left            =   3780
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   855
      Width           =   2895
      Begin VB.CheckBox XNODEM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰ «’‰«› »œÊ‰ ÿ·»Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   765
         Width           =   2445
      End
      Begin VB.CheckBox XISSTOP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰ «’‰«› „ Êﬁ›…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   495
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   450
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.CheckBox XNOMOVE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«’‰«› ·Ì” ·Â« «Ï Õ—ﬂ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   180
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      Height          =   690
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   675
      Width           =   3390
      Begin VB.CommandButton CMD_DELITEM 
         BackColor       =   &H008080FF&
         Caption         =   "Õ–› «·«’‰«›"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   45
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   135
         Width           =   3300
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1350
      Width           =   3390
      Begin VB.CommandButton cmd_excel 
         Height          =   555
         Left            =   1170
         Picture         =   "DataItem_All.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1005
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   2205
         Picture         =   "DataItem_All.frx":25BF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "DataItem_All.frx":4AB1
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   135
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2085
      Left            =   6705
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   0
      Width           =   11805
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "»«—ﬂÊœ Œ«—ÃÏ „ﬂ——"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   135
         Width           =   1950
      End
      Begin VB.TextBox XBARCODE 
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
         Left            =   135
         MaxLength       =   15
         TabIndex        =   8
         Top             =   540
         Width           =   2445
      End
      Begin VB.TextBox xColor 
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
         Left            =   7065
         MaxLength       =   15
         TabIndex        =   3
         Top             =   900
         Width           =   1545
      End
      Begin VB.TextBox XITEM 
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
         Left            =   135
         MaxLength       =   15
         TabIndex        =   12
         Top             =   1305
         Width           =   1815
      End
      Begin VB.TextBox xScal 
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
         Left            =   8955
         MaxLength       =   15
         TabIndex        =   2
         Top             =   900
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
         Left            =   4095
         MaxLength       =   15
         TabIndex        =   7
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox xDesca 
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
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   922
         Width           =   5505
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   7065
         TabIndex        =   1
         Top             =   540
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
         Left            =   7065
         TabIndex        =   0
         Top             =   180
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
         Left            =   3330
         TabIndex        =   6
         Top             =   180
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XBRNCH 
         Height          =   315
         Left            =   7065
         TabIndex        =   4
         Top             =   1305
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XSECTION 
         Height          =   315
         Left            =   3465
         TabIndex        =   10
         Top             =   1305
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xage 
         Height          =   315
         Left            =   7065
         TabIndex        =   5
         Top             =   1665
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483643
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xsex 
         Height          =   315
         Left            =   3465
         TabIndex        =   11
         Top             =   1665
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483643
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "›∆… ⁄„—Ì… :"
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
         Left            =   10710
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1665
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "«·‰Ê⁄ :"
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
         Index           =   7
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1665
         Width           =   555
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "»«—ﬂÊœ Œ«—ÃÏ-GS1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   585
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·ﬁ”„ : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5715
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1305
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·›—⁄ : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   10635
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1305
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "»«—ﬂÊœ "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1350
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "„ﬁ«” - ·Ê‰ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   10635
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   945
         Width           =   1035
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ „ÊœÌ· :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5715
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   630
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ⁄‰ ’‰› :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   4
         Left            =   5715
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„Ã„Ê⁄… :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   10635
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·„’‰⁄ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10635
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   255
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ê”„ :"
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
         Index           =   3
         Left            =   5715
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   225
         Width           =   780
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   13
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
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   3105
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   2430
      Top             =   1305
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
      Left            =   4080
      Top             =   1440
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
      Left            =   3495
      Top             =   1440
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
      Left            =   3195
      Top             =   1440
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
      Left            =   3795
      Top             =   1440
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3195
      Top             =   1395
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   25
      Top             =   10560
      Width           =   20400
      _ExtentX        =   35983
      _ExtentY        =   265
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc DATA7 
      Height          =   330
      Left            =   3195
      Top             =   1395
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
   Begin MSAdodcLib.Adodc DATA8 
      Height          =   330
      Left            =   3195
      Top             =   1395
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
      Height          =   7125
      Left            =   135
      TabIndex        =   33
      Top             =   2115
      Width           =   18375
      _cx             =   32411
      _cy             =   12568
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Cols            =   22
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   600
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
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
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
   Begin MSAdodcLib.Adodc DATA9 
      Height          =   330
      Left            =   3195
      Top             =   1395
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
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   330
      Left            =   3195
      Top             =   1395
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
   Begin VB.Label XMODEL 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3735
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1575
      Width           =   2085
   End
End
Attribute VB_Name = "DataItem_All"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myform As Variant, sM_Fact As String
Dim oSearch As New Search3
Dim con As New ADODB.Connection
Dim StoreTable As New ADODB.Recordset
Private Sub CMD_DELITEM_Click()
On Error GoTo myError
With grid1
    If MsgBox(" Õ–› «·„ÊœÌ·«  ›Ï «· ﬁ—Ì— ", vbYesNo + vbDefaultButton2) = vbYes Then
        prog1.Visible = True
        prog1.Value = 0
        prog1.Min = 0
        prog1.Max = .Rows - 1
        For nRow = 1 To .Rows - 1
            XMODEL.Caption = Val(.TextMatrix(nRow, 0))
            If GetDesca("SELECT ITEM FROM FILE1_11_ALL WHERE  ITEM = " & Val(.TextMatrix(nRow, 0)), con) = "" Then
                con.Execute " DELETE FROM FILE6_50 WHERE ITEM = " & Val(.TextMatrix(nRow, 0))
                con.Execute " DELETE FROM FILE6_51 WHERE ITEM = " & Val(.TextMatrix(nRow, 0))
                con.Execute " DELETE FROM FILE6_52 WHERE ITEM = " & Val(.TextMatrix(nRow, 0))
                con.Execute " DELETE FROM FILE1_10 WHERE ITEM = " & Val(.TextMatrix(nRow, 0))
            End If
            prog1.Value = nRow
        Next nRow
        Inform "  „ Õ–› «·„ÊœÌ·«  "
    End If
End With
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
    openCon con
    CMD_DELITEM.Visible = bSupermode
    StoreTable.Open "SELECT * FROM FILE0_40 WHERE isstop = 0 ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
    
    
    Set DATA3.Recordset = myRecordSet("Select Code,DescA From File1_50 ORDER BY DESCA", con)
    Set xGroup.RowSource = DATA3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    Set DATA4.Recordset = myRecordSet("Select mosm ,descA From mosm ORDER BY date DESC ", con)
    Set xMosm.RowSource = DATA4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm
    
    Set DATA5.Recordset = myRecordSet("Select code ,desca From fact ORDER BY code ", con)
    Set xFact.RowSource = DATA5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    xFact.BoundText = sM_Fact
    
    
    Set DATA7.Recordset = myRecordSet("Select Code,DescA From STORE_BR ORDER BY CODE ", con)
    Set XBRNCH.RowSource = DATA7
    XBRNCH.ListField = "Desca"
    XBRNCH.BoundColumn = "Code"
    
    Set data8.Recordset = myRecordSet("Select Code,DescA From FILE1_10SC ", con)
    Set XSECTION.RowSource = data8
    XSECTION.ListField = "Desca"
    XSECTION.BoundColumn = "Code"
    
    Set data9.Recordset = mycmd("Select Code,DescA From FILE0_81 ", con)
    Set xsex.RowSource = data9
    xsex.ListField = "Desca"
    xsex.BoundColumn = "Code"
    
    Set data10.Recordset = mycmd("Select Code,DescA From FILE0_82 ", con)
    Set xage.RowSource = data10
    xage.ListField = "Desca"
    xage.BoundColumn = "Code"
    
    Set grid1.DataSource = DATA6
   
    fixGrd
    If xFact.BoundText <> "" Then cmdGo_Click
End Sub
Private Sub myload()
    Dim cWhere As String, cFiled1 As String, cString  As String
    Dim aFields(21)
    aFields(0) = "FILE1_10.ITEM"
    aFields(1) = "FACT.DESCA AS FACT_DESCA"
    aFields(2) = "FILE1_10.MOSM"
    aFields(3) = "FILE1_10.MODELFACT0"
    aFields(4) = "FILE1_10.desca"
    aFields(5) = "FILE1_10.SCAL"
    aFields(6) = "FILE1_10.COLOR"
    aFields(7) = "FILE1_10.PRICE"
    aFields(8) = "FILE1_10.COST"
    aFields(9) = "FILE1_10.COST2"
    aFields(10) = "FILE1_10.BARCODE"
    
    aFields(11) = "FILE1_10.BARCODE2"
    aFields(12) = "FILE1_10.BARCODE13"
    aFields(13) = "FILE1_10.BARCODE_GS1"
    
    aFields(14) = "FILE1_10.DESCE"
    aFields(15) = "FILE1_10.COLOR_E"
    aFields(16) = "CASE WHEN PRICE_P = 0 THEN PRICE ELSE  FILE1_10.PRICE_P END  AS PRICE_P"
    aFields(17) = "FILE1_50.DESCA"
    aFields(18) = "FILE1_10.MODEL"
    
    aFields(19) = "FILE1_10SC.DESCA"
    aFields(20) = "FILE0_81.DESCA AS [SEX]"
    aFields(21) = "FILE0_82.DESCA AS [AGE]"
       
   '                        0               1                           2               3                   4                       5                   6               7           8                       9               10      11          12          13          14      15                      16                                                                  17  18
    'cString = " SELECT  FILE1_10.ITEM , FACT.DESCA AS FACTDESCA, FILE1_10.MOSM, FILE1_10.MODELFACT0, FILE1_10.desca AS DESCA,  FILE1_10.SCAL ,  FILE1_10.COLOR, FILE1_10.PRICE , FILE1_10.PRICE2 , FILE1_10.PRICE_2  , BARCODE , BARCODE2 , BARCODE13 , BARCODE_GS1 , DESCE , COLOR_E ,CASE WHEN PRICE_P = 0 THEN PRICE ELSE  FILE1_10.PRICE_P END  AS PRICE_P ,file1_50.desca, MODEL,FILE1_10SC.DESCA "
    
    cString = "SELECT " & ArString(aFields) & _
                " FROM  FILE1_10 INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE " & _
                " LEFT JOIN FILE1_50 ON FILE1_50.CODE = FILE1_10.[GROUP]" & _
                " LEFT JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE" & _
                " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX" & _
                " LEFT JOIN FILE0_82 ON FILE1_10.MODELAGE = FILE0_82.CODE"
    
    xmodelfact.text = DelZero(xmodelfact.text)
    
    If xMosm.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If Trim(xmodelfact.text) <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELFACT0 = " & MyParn(xmodelfact.text)
    If xGroup.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & "  FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & "  FILE1_10.FACT = " & MyParn(xFact.BoundText)
    
    If XSECTION.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & "  FILE1_10.SECTION = " & Val(XSECTION.BoundText)
    
    If XBRNCH.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ITEM IN (SELECT ITEM FROM FILE1_11_ALL WHERE STORE = " & MyParn(XBRNCH.BoundText) & ")"
    
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & "  FILE1_10.FACT = " & MyParn(xFact.BoundText)
    
    If xdesca.text <> "" Then cWhere = cWhere & Tr(cWhere) & " " & MyParnAnd(xdesca.text, "file1_10.desca")
    
    If xitem.text <> "" Then cWhere = cWhere & Tr(cWhere) & "  FILE1_10.item = " & Val(xitem.text)
    
    If xbarcode.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( FILE1_10.BARCODE = " & MyParn(xbarcode.text) & " OR  FILE1_10.BARCODE_GS1 = " & MyParn(xbarcode.text) & " ) "
    
    If Trim(xScal.text) <> "" Then cWhere = cWhere & Tr(cWhere) & "  FILE1_10.SCAL = " & MyParn(xScal.text)
    
    If Trim(xColor.text) <> "" Then cWhere = cWhere & Tr(cWhere) & "  " & MyParnAnd(xColor, "FILE1_10.COLOR")
    
    If Check1.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " BARCODE IN (SELECT BARCODE FROM QD_BARCODE )"
    
    If XNOMOVE.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " ITEM not IN (SELECT ITEM FROM FILE1_11) AND ITEM not IN (SELECT ITEM FROM FR1_11 ) "
    
    If XNODEM.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " ISNODEM = 0 "
    If xisstop.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " ISSTOP = 0 "
    
    If xage.MatchedWithList Then
         cWhere = cWhere & Tr(cWhere) & " MODELAGE =  " & xage.BoundText
    End If
    If xsex.MatchedWithList Then
         cWhere = cWhere & Tr(cWhere) & " MODELSEX =  " & xsex.BoundText
    End If
    
    If chkBarCode.Value = 1 Then
         cWhere = cWhere & Tr(cWhere) & " BARCODE_GS1 IS NULL"
    End If
    If cWhere <> "" Then cString = cString & " Where " & cWhere
    
    cString = cString & " ORDER BY FACT.DESCA , FILE1_10.MOSM, FILE1_10.MODELFACT0, FILE1_10.COLOR ,  FILE1_10.SCAL , PRICE_P"
    Set DATA6.Recordset = myRecordSet(cString, con)
    fixGrd
    CellPos 13, 0, grid1.Cols - 1
'    myload2
End Sub
Sub fixGrd()
With grid1
    
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .WordWrap = True
    
    .RowHeight(0) = 800
    
    .FrozenCols = 7
    
    
    .TextMatrix(0, 0) = "»«—ﬂÊœ"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .TextMatrix(0, 2) = "„Ê”„"
    .TextMatrix(0, 3) = "„ÊœÌ·"
    .TextMatrix(0, 4) = "«·’‰›"
    .TextMatrix(0, 5) = " „ﬁ«”"
    .TextMatrix(0, 6) = " «··Ê‰"
    .TextMatrix(0, 7) = " ”⁄— «·»Ì⁄"
    .TextMatrix(0, 8) = "”⁄— „’‰⁄"
    .TextMatrix(0, 9) = "”⁄— Ã„·…"
    
    .TextMatrix(0, 10) = " »«—ﬂÊœ Œ«—ÃÏ"
    .TextMatrix(0, 11) = " »«—ﬂÊœ „’‰⁄"
    .TextMatrix(0, 12) = " »«—ﬂÊœ 13"
    .TextMatrix(0, 13) = "BARCODE GS1"
    
    
    .TextMatrix(0, 14) = "Description"
    .TextMatrix(0, 15) = "COLOR"
    .TextMatrix(0, 16) = "”⁄— ÿ»«⁄… 2"
    .TextMatrix(0, 17) = "„Ã„Ê⁄…"
    .TextMatrix(0, 18) = "„Ã„Ê⁄…"
    .TextMatrix(0, 19) = "«·ﬁ”„"
    .TextMatrix(0, 20) = "«·‰Ê⁄"
    .TextMatrix(0, 21) = "«·›∆… «·⁄„—Ì…"
            
    .ColHidden(18) = True
    .ColWidth(0) = 900
    .ColWidth(1) = 2000
    .ColWidth(2) = 600
    .ColWidth(3) = 1500
    .ColWidth(4) = 4000
    .ColWidth(5) = 1000
    .ColWidth(6) = 2500
    .ColWidth(7) = 800
    .ColWidth(8) = 800
    .ColWidth(9) = 900
    .FrozenCols = 7
    .ColWidth(10) = 2200
    .ColWidth(11) = 1000
    .ColWidth(12) = 1700
    .ColWidth(13) = 1700
    .ColWidth(14) = 3000
    .ColWidth(15) = 1100
    .ColWidth(16) = 1000
    .ColWidth(17) = 2000
    .ColWidth(19) = 2000
    .ColWidth(20) = 2000
    .ColWidth(21) = 2000
        
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTString
    .ColDataType(4) = flexDTString
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    If .Rows > 1 Then .Row = 1
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Sub myProc()
If ActiveControl.Name = xFact.Name Then
    xFact.BoundText = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End If
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim cColor As String
On Error GoTo myError
With grid1
    If Col = 10 Then
        con.Execute " UPDATE FILE1_10 SET BARCODE = " & addstring(.TextMatrix(Row, 10)) & " WHERE ITEM = " & Val(.TextMatrix(Row, 0))
    
        For sbr = 3 To nCountBranch
            cBr = RetZero(sbr, 2)
            con.Execute "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) VALUES ( " & Val(.TextMatrix(Row, 0)) & " , " & addstring(cBr) & ")"
        Next sbr
        
        For sbr = 61 To nCountBranch_fr
            cBr = RetZero(sbr, 2)
            con.Execute "  INSERT INTO NEW_ITEM_FR (ITEM ,  BRANCH) VALUES ( " & Val(.TextMatrix(Row, 0)) & " , " & addstring(cBr) & ")"
        Next sbr
    ElseIf Col = 14 Then
        con.Execute " UPDATE FILE1_10 SET DESCE = " & addstring(.TextMatrix(Row, 14)) & " WHERE MODEL = " & MyParn(.TextMatrix(Row, 18)), nRec
        Inform "  „  ⁄œÌ· " & nRec
    ElseIf Col = 15 Then
        cColor = .TextMatrix(Row, 15)
        con.Execute " UPDATE FILE1_10 SET COLOR_E = " & addstring(.TextMatrix(Row, 15)) & " WHERE MODEL = " & MyParn(.TextMatrix(Row, 18)) & " AND COLOR = " & MyParn(.TextMatrix(Row, 6)), nRec
        Inform "  „  ⁄œÌ· " & nRec
    ElseIf Col = 4 Then
        con.Execute " UPDATE FILE1_10 ET DESCA = " & addstring(.TextMatrix(Row, 4)) & " WHERE MODEL = " & MyParn(.TextMatrix(Row, 18)), nRec
        Inform "  „  ⁄œÌ· " & nRec
    ElseIf Col = 6 Then
        con.Execute " UPDATE FILE1_10 SET COLOR = " & addstring(.TextMatrix(Row, 6)) & " WHERE COLOR_E = " & MyParn(.TextMatrix(Row, 15)) & " AND MODEL = " & MyParn(.TextMatrix(Row, 18)), nRec
        Inform "  „  ⁄œÌ· " & nRec
    ElseIf Col = 13 Then
        con.Execute " UPDATE FILE1_10 SET BARCODE_GS1 = " & addstring(.TextMatrix(Row, 13)) & " WHERE ITEM = " & .ValueMatrix(Row, 0)
    ElseIf Col = 7 Then
        con.Execute " UPDATE FILE1_10 SET PRICE = " & .ValueMatrix(Row, Col) & " WHERE ITEM = " & .ValueMatrix(Row, 0)
    ElseIf Col = 8 Then
        con.Execute " UPDATE FILE1_10 SET COST = " & .ValueMatrix(Row, Col) & " WHERE ITEM = " & .ValueMatrix(Row, 0)
    ElseIf Col = 9 Then
        con.Execute " UPDATE FILE1_10 SET COST2 = " & .ValueMatrix(Row, Col) & " WHERE ITEM = " & .ValueMatrix(Row, 0)
    End If
End With
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub grid1_DblClick()
    items.sModelItem = grid1.TextMatrix(grid1.Row, 0)
    items.Show 1
End Sub
Private Sub grid1_EnterCell()
With grid1
    If .Col = 7 Or .Col = 8 Or .Col = 9 Or .Col = 10 Or (.Col = 13 And bOpt10) Or .Col = 14 Or .Col = 15 Or .Col = 4 Or .Col = 6 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub

Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 13 Then
    If grid1.EditText = grid1.TextMatrix(Row, Col) Then
        Cancel = True
    ElseIf Len(grid1.EditText) < 13 And Trim(grid1.EditText) <> "" Then
        Cancel = True
    End If
End If
End Sub

Private Sub xFACT_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Set Generalarray(0) = Me
Generalarray(1) = "Select code ,DescA From FACT"
Generalarray(2) = "Order by DESCA"
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

End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 And ActiveControl.Name <> grid1.Name Then
 '   cmdGo_Click
'End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'If grid1.Col = 3 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    'If Col = 3 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
     CellPos KeyCode, grid1.Row, grid1.Col
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col = 13 Then
    If Row < grid1.Rows - 1 Then
        grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 13, 13)
        grid1.ShowCell grid1.Row, 13
    End If
ElseIf Col < grid1.Cols - 11 Then
    If Col < 4 Then
        grid1.Col = 4
    ElseIf Col < 6 Then
        grid1.Col = 6
    ElseIf Col < 6 Then
        grid1.Col = 6
    Else
        grid1.Col = Col + 1 + IIf(Col = 11, 1, 0)
    End If
    grid1.ShowCell grid1.Row, grid1.Col
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 0, 4)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, Col
End If
End Sub

