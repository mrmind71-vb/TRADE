VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form VsTitem_H 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„ «»⁄… „ÊœÌ·«  «·«„«‰« "
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   8010
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   1665
      Width           =   2715
   End
   Begin VB.Frame Frame2 
      Height          =   1725
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   45
      Width           =   3840
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   180
         Width           =   780
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "‰ «·„»Ì⁄«  :"
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
         Height          =   330
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   180
         Width           =   915
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «·„ÊœÌ·«  :"
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
         Height          =   330
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… „‘ —Ì«  «„«‰« "
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
         Height          =   330
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   555
         Width           =   1590
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… —’Ìœ «„«‰«  "
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
         Height          =   330
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   930
         Width           =   1590
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " ﬂ·›… „»Ì⁄«  «„«‰«  "
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
         Height          =   330
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1305
         Width           =   1590
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   555
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1305
         Width           =   1725
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   930
         Width           =   1725
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1710
      Width           =   3840
      Begin VB.CommandButton cmdPrint 
         Height          =   510
         Left            =   1305
         Picture         =   "VsTitem_H.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   90
         Picture         =   "VsTitem_H.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   510
         Left            =   2520
         Picture         =   "VsTitem_H.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "’Ê— «·„ÊœÌ·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   135
      Width           =   3570
      Begin ImgXCtrl6.ImgXCtrl ImgX1 
         Height          =   1995
         Left            =   90
         TabIndex        =   13
         Top             =   270
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   3519
         BackColor       =   16777215
         BorderStyle     =   4
         AutoZoom        =   -1  'True
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         AutoZoomType    =   3
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "mrvb71"
         LicenseRegCode  =   "íß“ªß•≤º∂´≠“±®ππ∂´µßZQEH-AOZOOOZT-EFLF6gI"
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   90
      Width           =   7440
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
         Left            =   135
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
         TabIndex        =   6
         Top             =   900
         Width           =   3915
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
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   4770
         TabIndex        =   4
         Top             =   945
         Width           =   1635
         _ExtentX        =   2884
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
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   315
         Left            =   2970
         TabIndex        =   10
         Top             =   540
         Width           =   3435
         _ExtentX        =   6059
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
      Begin VB.Label xModelNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   135
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ :"
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
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   630
         Width           =   600
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
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   585
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·»Ì«‰ :"
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
         Left            =   4095
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   990
         Width           =   540
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
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   990
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï  «—ÌŒ :"
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
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   270
         Width           =   870
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
   Begin VSFlex7Ctl.VSFlexGrid GRID1 
      Height          =   7665
      Left            =   45
      TabIndex        =   19
      Top             =   2610
      Width           =   14955
      _cx             =   26379
      _cy             =   13520
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   16777152
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
      Rows            =   2
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
      ScrollBars      =   2
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
   Begin MSAdodcLib.Adodc data20 
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
   Begin ComctlLib.ProgressBar xbar 
      Height          =   195
      Left            =   45
      TabIndex        =   29
      Top             =   2430
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "VsTitem_H"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearch As New Search3
Public cMyModel As String
Dim con As New ADODB.Connection
Private Sub cmd_dem_Click()
    grditemfrm1.Show
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    xCount.Caption = ""
    xTPurch.Caption = ""
    xTSalCOST.Caption = ""
    xTBal.Caption = ""
    myload
End Sub
Private Sub cmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = " »Ì«‰ ≈Ã„«·Ï  ⁄«„·«  „ÊœÌ·«  «·√„«‰«  ··„Ê—œ " & xSupp.Text
    cHead2 = GetDesca("SELECT PHONE1 FROM FILE4_10 WHERE CODE = " & MyParn(xSupp.BoundText))
    With grid1
    .ColHidden(6) = True
    .ColHidden(16) = True
    
    
    
    Load PrintGrd
    PrintGrd.doprint Me.grid1, 0.85, 2, cHead1, cHead2, , False, , 7
    PrintGrd.Show 1

    .ColHidden(6) = False
    .ColHidden(16) = False
    End With
End Sub
Private Sub Command1_Click()
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset, nBalance As Single
Dim ModelsTable As New ADODB.Recordset
Dim ScalTable As New ADODB.Recordset
Dim ColorTable As New ADODB.Recordset

Dim aHeader(2)
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

'cStr1 = " SELECT  FILE1_10.MODEL, FILE1_10.desca, FILE1_10.mosm , SUM(FILE1_11.[IN] - FILE1_11.OUT ) AS BALMODEL  FROM  FILE1_10  INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM "
'If xMosm.BoundText <> "" Then cStr1 = cStr1 & " WHERE FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
'If xGroup.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
'If xModel.Text <> "" Then cStr1 = cStr1 & " AND FILE1_10.[MODEL] = " & MyParn(xModel.Text)
'cStr1 = cStr1 & " GROUP BY FILE1_10.MODEL, FILE1_10.desca, FILE1_10.mosm ORDER BY FILE1_10.MODEL "
'
'ModelsTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText

Dim cModel As String
Dim TModelDem As Recordset
Dim nTotal As Double
Dim nCol As Double
With temptable
For nRow = 3 To grid1.Rows - 1
    If Val(grid1.TextMatrix(nRow, 10)) <> 0 Then
        cModel = grid1.TextMatrix(nRow, 0)
        cStr1 = " SELECT FILE1_10.MODELfact, FILE1_10.C_SCAL, FILE1_10.SCAL, SUM(FILE1_11.[IN] - FILE1_11.OUT ) AS T_BAL Into BAL_Rep1 FROM  FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM  WHERE FILE1_10.MODEL = " & MyParn(cModel)
        If IsDate(xDate1.Text) Then cStr1 = cStr1 & " AND FILE1_11.[date] <= " & DateSq(xDate1.Text)
        cStr1 = cStr1 & "GROUP BY FILE1_10.MODELfact, FILE1_10.C_SCAL, FILE1_10.SCAL"
        DelTable
        con.Execute cStr1
        
        If ScalTable.State = adStateOpen Then ScalTable.Close
        cStr1 = "SELECT cost , C_SCAL , SCAL FROM FILE1_10 WHERE MODEL =  " & MyParn(cModel) & " GROUP BY cost , FILE1_10.C_SCAL , SCAL ORDER BY C_SCAL "
        ScalTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
        .AddNew
        !str17 = cModel
        !str16 = grid1.TextMatrix(nRow, 4)
        cGr = GetDesca("SELECT [GROUP] FROM FILE1_10 WHERE MODEL = " & MyParn(cModel))
        If cGr <> "" Then
            !str13 = TurnValue(GetDesca("select descaA from file1_50 where code = " & MyParn(cGr)), "", Null)
        End If
        !str18 = ArbString(grid1.TextMatrix(nRow, 5))
        !str15 = " ›’Ì·Ï —’Ìœ «·„ÊœÌ·« "
        nTotal = 0
        ntcost = 0
        ScalTable.MoveFirst
        nCol = 0
        Do While Not ScalTable.EOF
            nCol = nCol + 1
            .Fields("STR" & nCol) = ScalTable!scal
            .Fields("VAL" & nCol) = Val(GetDesca("SELECT T_BAL FROM BAL_Rep1 WHERE SCAL = " & MyParn(ScalTable!scal)) & "")
            nTotal = nTotal + Val((.Fields("VAL" & nCol)) & "")
            ntcost = ntcost + (Val((.Fields("VAL" & nCol)) & "") * ScalTable!cost)
            ScalTable.MoveNext
        Loop
        !val13 = nTotal
        !VAL20 = ntcost
        .Update
    End If
Next nRow
End With
contemp.BeginTrans
contemp.CommitTrans
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    main.Report1.ReportFileName = App.Path & "\Reports\R_balscal.rpt"
    main.Report1.DataFiles(0) = tempFile
    main.Report1.Action = 1
End If

temptable.Close


Set temptable = Nothing
Set sourcetable = Nothing

End Sub

Private Sub Form_Load()
    openCon con
    xDate1.Text = Format(Date, "dd-mm-yyyy")
    
    
    Dim ImgXTmp As ImgX
    Set ImgXTmp = New ImgX
'    ImgXTmp.Import.FromMemoryFile LoadResData("BACKGROUND", "CUSTOM")
'    Set ImgX1.BackgroundImage = ImgXTmp.image
    Set ImgXTmp = Nothing
    
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "Select Code,DescA From File4_10 order by Desca"
    Set xSupp.RowSource = DATA2
    xSupp.ListField = "Desca"
    xSupp.BoundColumn = "Code"
    
    DATA4.ConnectionString = strCon
    DATA4.RecordSource = "Select mosm ,descA From mosm ORDER BY date "
    Set xMosm.RowSource = DATA4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm
    
    Set grid1.DataSource = DATA6
    DATA6.ConnectionString = strCon
    
    
    grid1.Rows = 2
    grid1.Cols = 17
    Fixgrd
End Sub
Private Sub myload()
    Dim cwhere As String
    
    cField7 = myiif("type = '2' ", "([IN] - [OUT]) ") & " AS t_purch, "
    cField8 = myiif("type = '7' ", "([out] - [in]) ") & " AS t_ret  , "
    cField9 = myiif("type = '6' ", "([out] - [in]) ") & " AS t_sal  , "
    cField10 = "SUM ([IN] - [OUT])  AS t_sal  , "
    
    cField11 = myiif("type = '2' OR type = '7' ", "(([IN]-[OUT]) * FILE1_10.COST ) ") & " AS V_purch, "
    cField12 = myiif("type = '6' ", "([out] * FILE1_10.COST  ) ") & " AS V_sal  , "
    cField13 = "SUM (([IN] - [OUT]) * FILE1_10.COST  )   AS V_sal   "
    
'                           0               1               2               3           4                   5                                   6
    cStr1 = " SELECT   FILE1_10.MODELNO , FILE1_10.MOSM, FILE1_10.FACT, FACT.DESCA, FILE1_10.MODELFACT, FILE1_10.desca AS itemdesca  , MIN(FILE1_10.COST) AS cost , " & _
            cField7 & cField8 & cField9 & cField10 & cField11 & cField12 & cField13 & " , ' ' AS N14 , FILE1_10.OKAZ , dbo.sum_str_modelno(file1_10.modelno) AS N14 " & _
            " FROM   FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE INNER JOIN ITEM_ONEST ON FILE1_10.ITEM = ITEM_ONEST.ITEM WHERE FILE1_10.ITEM IS NOT NULL "
    If xMosm.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xSupp.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If IsDate(xDate1.Text) Then cStr1 = cStr1 & " AND date <= " & DateSq(xDate1.Text)
    cStr1 = cStr1 & " GROUP BY FILE1_10.MODELNO , FILE1_10.MOSM, FILE1_10.FACT, FACT.DESCA, FILE1_10.MODELFACT, FILE1_10.desca , FILE1_10.OKAZ ORDER BY   file1_10.MODELNO , FILE1_10.MOSM, FILE1_10.FACT, FILE1_10.MODELFACT "
    DATA6.RecordSource = cStr1
    DATA6.Refresh
    Fixgrd
End Sub
Sub Fixgrd()
    Dim nPurch As Double
    Dim nBal As Double
    Dim nsAL As Double
 
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 2
    .RowHidden(1) = True
    .RowHeight(0) = 1000
    .WordWrap = True
    .Cols = 17
    .ColHidden(0) = True
    
    .TextMatrix(0, 0) = "„Ê—œ"
    .TextMatrix(0, 1) = "„Ê”„"
    .TextMatrix(0, 2) = "ﬂÊœ"
    .TextMatrix(0, 3) = "«·„’‰⁄"
    .TextMatrix(0, 4) = "„ÊœÌ· "
    .TextMatrix(0, 5) = "«·’‰›"
    .TextMatrix(0, 6) = "”⁄— «·„’‰⁄"
    
    .TextMatrix(0, 7) = "Ã. „‘ —Ì«  "
    .TextMatrix(0, 8) = "„— Ã⁄« "
    .TextMatrix(0, 9) = "„»Ì⁄« "
    .TextMatrix(0, 10) = "—’Ìœ"
    
    .TextMatrix(0, 11) = "ﬁÌ„… „‘ —Ì«  «„«‰« "
    .TextMatrix(0, 12) = "ﬁÌ„…  ﬂ·›… „»Ì⁄«  «„«‰« "
    .TextMatrix(0, 13) = "ﬁÌ„… —’Ìœ √„«‰« "
    
    .TextMatrix(0, 14) = "‰”»… „»Ì⁄« "
    .TextMatrix(0, 15) = "‰”»… «Êﬂ«“ÊÌ‰"
    
    .TextMatrix(0, 16) = "—ÌÊ‰"
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 700
    .ColWidth(2) = 700
    .ColWidth(3) = 1700
    .ColWidth(4) = 700
    .ColWidth(5) = 2600
    
    .ColWidth(6) = 600
    .ColWidth(7) = 600
    .ColWidth(8) = 600
    .ColWidth(9) = 600
    .ColWidth(10) = 600
    
    .ColWidth(11) = 900
    .ColWidth(12) = 900
    .ColWidth(13) = 900
    
    .ColWidth(14) = 700
    .ColWidth(15) = 700
    
    .ColWidth(16) = 700
    
    .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = 4
    
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTString
    .ColDataType(4) = flexDTString
    
    For nCol = 5 To .Cols - 2
        .ColDataType(nCol) = flexDTDouble
    Next nCol
    .Editable = flexEDNone
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
    .Subtotal flexSTClear
    .WordWrap = True
        
    For nRow = 1 To .Rows - 1
        nPurch = Val(.TextMatrix(nRow, 7)) - Val(.TextMatrix(nRow, 8))
        nRate = 0
        If nPurch <> 0 Then nRate = Val(.TextMatrix(nRow, 9)) / nPurch * 100
        .TextMatrix(nRow, 14) = Round(nRate, 2)
    Next nRow
    For nCol = 7 To 13
        .Subtotal flexSTSum, -1, nCol, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    Next nCol
    If .Rows > 2 Then
        nRow = 2
        nPurch = Val(.TextMatrix(nRow, 7)) - Val(.TextMatrix(nRow, 8))
        nRate = 0
        If nPurch <> 0 Then nRate = Val(.TextMatrix(nRow, 9)) / nPurch * 100
        .TextMatrix(nRow, 14) = Round(nRate, 2)
        xCount.Caption = .Rows - 3
        xTPurch.Caption = .TextMatrix(2, 11)
        xTSalCOST.Caption = .TextMatrix(2, 12)
        xTBal.Caption = .TextMatrix(2, 13)
        xRate.Caption = Round(nRate, 2)
    End If
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub grid1_DBLClick()
    Dim cString As String
    
    If grid1.Col <= 3 Then
        
        Load ModelMove
        ModelMove.xitem.Text = xmodelno.Caption
        ModelMove.Show
    
    ElseIf grid1.Col = 7 Or grid1.Col = 8 Then
        
        InvModel.pModelNo = xmodelno.Caption
        InvModel.Show
    
    ElseIf grid1.Col = 9 And Val(grid1.TextMatrix(grid1.Row, 9)) <> 0 Then
        
        cString = "SELECT SUM(QUANT) AS sumOfQuant,FILE1_10.COLOR,FILE1_10.SCAL" & _
                  " FROM FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM" & _
                  " INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO"
'        If IsDate(xDate1.Text) Then cString = cString & turn(cString) & "FILE6_20H.DATE <= " & DateSq(xDate1.Text)
        cString = cString & turn(cString) & " FILE1_10.modelno = " & MyParn(xmodelno.Caption)
        cString = cString & " GROUP BY FILE1_10.COLOR,FILE1_10.SCAL"
        grdModelfrm.pString = cString
        grdModelfrm.pFieldValue = "sumofQuant"
        grdModelfrm.pModelNo = xmodelno.Caption
        grdModelfrm.Show 1
    
    ElseIf grid1.Col = 10 And Val(grid1.TextMatrix(grid1.Row, 10)) <> 0 Then
        
        cString = "SELECT SUM([IN]-[OUT]) AS sumOfQuant,FILE1_10.COLOR,FILE1_10.SCAL" & _
                  " FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM"
'        If IsDate(xDate1.Text) Then cString = cString & turn(cString) & "FILE1_11.DATE <= " & DateSq(xDate1.Text)
        cString = cString & turn(cString) & " FILE1_10.modelno = " & MyParn(xmodelno.Caption)
        cString = cString & " GROUP BY FILE1_10.COLOR,FILE1_10.SCAL"
        grdModelfrm.pString = cString
        grdModelfrm.pFieldValue = "sumofQuant"
        grdModelfrm.pModelNo = grid1.TextMatrix(grid1.Row, 0)
        grdModelfrm.Show 1
    End If
End Sub
Private Sub grid1_EnterCell()
    xmodelno.Caption = grid1.TextMatrix(grid1.Row, 0)
    LoadPhoto xmodelno.Caption
'   GRID1.ToolTipText = GetDesca("SELECT DESCA FROM FACT WHERE CODE = " & MyParn(GRID1.TextMatrix(GRID1.Row, 1)))
    If grid1.Col = 15 Then
        grid1.Editable = flexEDKbdMouse
    Else
        grid1.Editable = flexEDNone
    End If
End Sub
Private Sub ImgX1_Click()
If validPhoto(retFile(grid1.TextMatrix(grid1.Row, 0))) Then
    Dim oMaxPhoto As New maxPhoto
    oMaxPhoto.cFile = retFile(grid1.TextMatrix(grid1.Row, 0))
    oMaxPhoto.cDesca = grid1.TextMatrix(grid1.Row, 4)
    oMaxPhoto.Show 1
End If
End Sub
Private Sub LoadPhoto(sModelNo As String)
On Error GoTo myerror
ImgX1.Images.Clear
ImgX1.Import.FromFile retFile(sModelNo)
Exit Sub
myerror:
Err.Clear
End Sub
Private Sub xSupp_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then SuppLookupAll Me, oSearch
End Sub
Sub myProc()
If ActiveControl.Name = xSupp.Name Then
    xSupp.BoundText = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    oSearch.Hide
End If
End Sub
Private Sub xsupp_Validate(Cancel As Boolean)
If xSupp.BoundText = "" Then Exit Sub
If Not xSupp.MatchedWithList Then
    cFact = GetDesca("SELECT CODE FROM FILE4_10 WHERE SUBCODE = " & MyParn(xSupp.BoundText))
    xSupp.BoundText = cFact
    If xSupp.MatchedWithList Then Exit Sub
End If
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
    If Not .IsSubtotal(Row) Then
        If .TextMatrix(Row, 15) <> "" Then
            DoFixPrice2 grid1.TextMatrix(grid1.Row, 0), Val(grid1.TextMatrix(grid1.Row, 15)), grid1.Row
            con.Execute " UPDATE FILE1_10H SET OKAZ          = " & Val(nOkaz) & " WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
        End If
    End If
    
    If .IsSubtotal(Row) Then
        If MsgBox("‰”»… √Êﬂ«“ÌÊ‰ ·ﬂ· «·„ÊœÌ·« ", vbYesNo + vbDefaultButton2) = vbYes Then
            xbar.Max = .Rows - 1
            xbar.Min = 0
            xbar.Value = 0
            For nRow = 2 To .Rows - 1
                If .TextMatrix(nRow, 15) <> Val(.TextMatrix(Row, 15)) Then
                   .TextMatrix(nRow, 15) = Val(.TextMatrix(Row, 15))
                    DoFixPrice2 grid1.TextMatrix(nRow, 0), Val(.TextMatrix(Row, 15)), nRow
                    con.Execute " UPDATE FILE1_10H SET OKAZ          = " & Val(nOkaz) & " WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
                    .Select nRow, 0, nRow, .Cols - 1
                End If
                xbar.Value = nRow
            Next nRow
        End If
    End If
End With
End Sub
Sub DoFixPrice2(cModel, nOkaz, nRow)
    Dim SubItemTable As New ADODB.Recordset
    Dim nRem As Double, nPrice As Double
    Dim nFPrice As Double
    cStr1 = " SELECT * FROM FILE1_10  WHERE MODELNO = " & MyParn(cModel)
    SubItemTable.Open cStr1, con, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SubItemTable.MoveFirst
        Do While Not SubItemTable.EOF
            If nOkaz > 0 Then
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
            Else
                nPrice = SubItemTable!price
            End If
            If nFPrice = 0 Then nFPrice = nPrice
            SubItemTable!PRICE2 = nPrice
            SubItemTable!okaz = nOkaz
            SubItemTable.Update
            SubItemTable.MoveNext
        Loop
End Sub


Private Sub DelTable()
On Error Resume Next
con.Execute "DROP TABLE BAL_Rep1 "
Err.Clear
End Sub

