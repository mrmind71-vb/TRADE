VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form OrderOnline_New 
   Caption         =   "ÿ·»Ì«  «Ê‰ ·«Ì‰"
   ClientHeight    =   11055
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   20370
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
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin Threed.SSCommand CMD_SEND 
      Height          =   1005
      Left            =   16110
      TabIndex        =   27
      Top             =   6885
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   1773
      _Version        =   196610
      BackColor       =   16761024
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "OrderOnline_New.frx":0000
      Caption         =   " —ÕÌ· ·›« Ê—… „»Ì⁄« "
      Alignment       =   8
      PictureAlignment=   6
   End
   Begin VB.Frame Frame1 
      Height          =   1440
      Left            =   5175
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   13485
      Begin VB.CheckBox RET_NOTSHIP 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ÿ·»Ì«  „—›Ê÷… Ê ·„ Ì „ «” ·«„Â« „‰ «·‘Õ‰"
         Height          =   690
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   720
         Width           =   1545
      End
      Begin VB.TextBox xphone 
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
         Left            =   10215
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   990
         Width           =   1995
      End
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   135
         Picture         =   "OrderOnline_New.frx":2808
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   855
         Width           =   1455
      End
      Begin VB.TextBox XDOC_NO 
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
         Left            =   10845
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   585
         Width           =   1365
      End
      Begin VB.CommandButton cmdGo 
         Height          =   600
         Left            =   135
         Picture         =   "OrderOnline_New.frx":4C32
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "⁄—÷"
         Top             =   225
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Height          =   600
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   720
         Width           =   3435
         Begin VB.OptionButton XSALES 
            Alignment       =   1  'Right Justify
            Caption         =   "«·ﬂ·"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   135
            Width           =   645
         End
         Begin VB.OptionButton XSALES 
            Alignment       =   1  'Right Justify
            Caption         =   "„»«⁄…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   135
            Width           =   780
         End
         Begin VB.OptionButton XSALES 
            Alignment       =   1  'Right Justify
            Caption         =   "€Ì— „»«⁄…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   135
            Value           =   -1  'True
            Width           =   1140
         End
      End
      Begin VB.Frame Frame2 
         Height          =   600
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   135
         Width           =   5010
         Begin VB.OptionButton XISSEND 
            Alignment       =   1  'Right Justify
            Caption         =   "ÿ·»Ì«  „·€«…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   180
            Width           =   1500
         End
         Begin VB.OptionButton XISSEND 
            Alignment       =   1  'Right Justify
            Caption         =   "«·ﬂ·"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1710
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   180
            Width           =   735
         End
         Begin VB.OptionButton XISSEND 
            Alignment       =   1  'Right Justify
            Caption         =   "„—Õ·… "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   2430
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   180
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton XISSEND 
            Alignment       =   1  'Right Justify
            Caption         =   "·„  —Õ· ··›—⁄"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   3375
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   180
            Width           =   1500
         End
      End
      Begin VB.TextBox xdate2 
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
         Left            =   7785
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   1545
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
         Left            =   10845
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo xpay 
         Height          =   330
         Left            =   6705
         TabIndex        =   31
         Top             =   585
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   330
         Left            =   6705
         TabIndex        =   33
         Top             =   990
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «· ·Ì›Ê‰"
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
         Left            =   12285
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "«·›—⁄ : "
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9360
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   900
         Width           =   930
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "«·”œ«œ "
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9450
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   540
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ ÿ·»Ì… : "
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
         Left            =   12285
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï  «—ÌŒ :"
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
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ :"
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
         Left            =   12330
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   765
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10725
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data3 
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
   Begin MSAdodcLib.Adodc data1 
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
   Begin MSAdodcLib.Adodc data5 
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
   Begin MSAdodcLib.Adodc DATA7 
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
   Begin MSAdodcLib.Adodc Adodc1 
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
      Caption         =   "data7"
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
   Begin MSAdodcLib.Adodc Adodc2 
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
   Begin VSFlex7Ctl.VSFlexGrid GRID2 
      Bindings        =   "OrderOnline_New.frx":7124
      Height          =   3285
      Left            =   -45
      TabIndex        =   14
      Top             =   6930
      Width           =   18510
      _cx             =   32650
      _cy             =   5794
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16776960
      ForeColorSel    =   64
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
      Rows            =   1
      Cols            =   10
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
   Begin MSComDlg.CommonDialog Common1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA4 
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
   Begin Threed.SSCommand cmdExcel 
      Height          =   510
      Left            =   90
      TabIndex        =   38
      Top             =   990
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   900
      _Version        =   196610
      ForeColor       =   0
      BackColor       =   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "OrderOnline_New.frx":7138
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin VB.Frame Frame4 
      Height          =   1005
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   5055
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "⁄—÷ —’Ìœ ÿ· «·›—Ê⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   180
         Width           =   2175
      End
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   45
         Picture         =   "OrderOnline_New.frx":9663
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   1185
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   240
         Left            =   90
         TabIndex        =   13
         Top             =   720
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   423
         _Version        =   327682
         Appearance      =   1
      End
      Begin Threed.SSCommand cmd_addexel 
         Height          =   525
         Left            =   1260
         TabIndex        =   23
         Top             =   135
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   926
         _Version        =   196610
         Font3D          =   3
         CaptionStyle    =   1
         ForeColor       =   64
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«÷«›… „‰ «ﬂ”Ì·"
         ButtonStyle     =   2
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "OrderOnline_New.frx":BACF
      Height          =   4860
      Left            =   90
      TabIndex        =   1
      Top             =   1530
      Width           =   18555
      _cx             =   32729
      _cy             =   8572
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16776960
      ForeColorSel    =   64
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
      Rows            =   1
      Cols            =   10
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
   Begin VB.Label XBRANCH 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1710
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   1035
      Width           =   3390
   End
   Begin VB.Label XTIME 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12375
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   6480
      Width           =   2895
   End
   Begin VB.Label XUSER 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15435
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   6480
      Width           =   3165
   End
   Begin VB.Label XNOTES 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   6480
      Width           =   12120
   End
End
Attribute VB_Name = "OrderOnline_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearch As New Search3, lIsBoxOnline As Boolean
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con_SERVER As New ADODB.Connection
Dim con_MyShop As New ADODB.Connection, oSearchItem As New Search3
Private Sub cmd_addexel_Click()
    AddFromExel
    Inform "  „ «÷«›… «·ÿ·»Ì«  "
    myload
End Sub
Private Sub CMD_PRINT_Click()
    doprint_day
End Sub
Private Sub CMD_SEND_Click()
    Dim cDocSalPost As String
    If cBranch = "00" Then Exit Sub
    If grid1.TextMatrix(grid1.Row, 16) = "" Then
        If MsgBox(" —ÕÌ· ·›« Ê—… „»Ì⁄«  ", vbYesNo + vbDefaultButton2) = vbYes Then
            cDocSalPost = myreplace_Sales()
            If cDocSalPost <> "" Then
                MsgBox " „  —ÕÌ· ·»Ê‰ „»Ì⁄«  —ﬁ„ " & cDocSalPost
            Else
                MsgBox "·„ Ì „ ⁄„· „” ‰œ «·„»Ì⁄«   "
            End If
        End If
        myload
        GRID2.Rows = 2
    End If
End Sub

Private Sub cmdExcel_Click()
ToFileExelNew grid1, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption)
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
'    On Error GoTo myerror
     
    openCon con_SERVER, LoadConString_OnLine
    If cBranch <> "00" Then openCon con_MyShop
    
    Set grid1.DataSource = data1
    Set GRID2.DataSource = DATA2
    
    Frame2.Visible = (cBranch = "00")
    Check1.Visible = (cBranch = "00")
    cmd_addexel.Visible = (cBranch = "00")
    
    Set DATA3.Recordset = myRecordSet("SELECT Payment_Method FROM file6_90h GROUP BY Payment_Method ", con_SERVER)
    Set xPay.RowSource = DATA3
    xPay.ListField = "Payment_Method"
    xPay.BoundColumn = "Payment_Method"
    
    Set DATA4.Recordset = myRecordSet("SELECT CODE , DESCA FROM FILE0_40 WHERE online =  1 ORDER BY CODE ", con_SERVER)
    Set xStore.RowSource = DATA4
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
        
    If GetBoolean("SELECT ONLINE FROM FILE0_40 WHERE CODE = " & MyParn(cBranchStore), con_SERVER) Then
        lIsBoxOnline = True
        If cBranchBox <> "" Then
            lIsBoxOnline = GetBoolean("SELECT ISBOXONLINE FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con_SERVER)
        End If
    End If
    
    CMD_SEND.Visible = (cBranch <> "00") And lIsBoxOnline And cManBox <> "" And Not lSupperVisor
    If cBranch <> "00" Then GRID2.Width = GRID2.Width - 3000
    grid1.Rows = 1
    grid1.Cols = 9
    FIXGRID
    XBRANCH.Caption = GetDesca("SELECT DESCA FROM BRANCH WHERE CODE = " & MyParn(cBranch), con_SERVER)
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub
Private Sub myload()
'On Error GoTo myerror
Dim I As Double
Dim cString  As String, cStr2 As String
Dim cF1 As String, cF2 As String, cF3 As String, cF4 As String, cF5 As String
Dim cSubTotal As String

With grid1
    cSubTotal = "(SELECT SUM(TOTAL) FROM FILE6_90 WHERE FILE6_90.DOC_NO = FILE6_90H.DOC_NO )    "
    cTotal = "(SELECT SUM(TOTAL) FROM FILE6_90 WHERE FILE6_90.DOC_NO = FILE6_90H.DOC_NO )  - DISCOUNT   "
    
'                           0     1     2     3        4            5           6                7        8                  9           10              11            12       13          14           15                 16                      17         18           19      20              21
    cString = " SELECT DOC_NO, DATE, NAME, PHONE, Shipping_City, STREET , " & cSubTotal & " , DISCOUNT, DISCOUNT_CODE, SHIPPING , " & cTotal & " , Payment_Method , STORE , SEND_USER ,  SEND_TIME , NOTES_ORDER  , SUBSTRING(SALES_DOC,7,20) , SALES_DATE , SHIP_NO  , NOTES , DelOrder_Date , DelOrder_Date2  FROM FILE6_90H WHERE  DOC_NO IS NOT NULL "
    If xDoc_No.text <> "" Then cString = cString & " AND [DOC_NO] = " & MyParn(xDoc_No.text)
    If xphone.text <> "" Then cString = cString & " AND [phone] = " & MyParn(xphone.text)
    If xPay.BoundText <> "" Then cString = cString & " AND [Payment_Method] = " & MyParn(xPay.text)
    If xStore.BoundText <> "" Then cString = cString & " AND [STORE] = " & MyParn(xStore.BoundText)
    If IsDate(xDate1.text) Then cString = cString & " AND [DATE] >= " & DateSq(xDate1.text)
    If IsDate(xDate2.text) Then cString = cString & " AND [DATE] <= " & DateSq(xDate2.text)
    If cBranch = "00" Then
        If xIssend(0).Value <> 0 Then cString = cString & " AND STORE IS NULL  and DelOrder_Date is null "
        If xIssend(1).Value <> 0 Then cString = cString & " AND STORE IS NOT  NULL AND DelOrder_Date IS NULL "
        If xIssend(3).Value <> 0 Then cString = cString & " AND DelOrder_Date IS NOT  NULL "
    Else
        cString = cString & " AND STORE = " & MyParn(cBranchStore)
    End If
    If xsales(0).Value <> 0 Then cString = cString & " AND SALES_DOC IS NULL "
    If xsales(1).Value <> 0 Then cString = cString & " AND SALES_DOC IS NOT NULL "
    
    If RET_NOTSHIP.Value <> 0 Then cString = cString & " AND DelOrder_Date IS NOT NULL AND DelOrder_Date2 IS NULL "
    cString = cString & " ORDER BY DOC_NO "
    Set data1.Recordset = myRecordSet(cString, con_SERVER)

End With
FIXGRID
grid1.Cell(flexcpAlignment, 0, 0, grid1.Rows - 1, grid1.Cols - 1) = 7
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub FIXGRID()
With grid1
    .RowHeight(0) = 1000
    .WordWrap = True
    .Cols = 22
    .FrozenCols = 3
    
    .TextMatrix(0, 0) = "—ﬁ„"
    .TextMatrix(0, 1) = "«· «—ÌÕ"
    .TextMatrix(0, 2) = "«·≈”„"
    .TextMatrix(0, 3) = " ·Ì›Ê‰"
    .TextMatrix(0, 4) = "«·„œÌ‰…"
    .TextMatrix(0, 5) = "⁄‰Ê«‰"
    
    .TextMatrix(0, 6) = "≈Ã„«·Ï «’‰«›"
    .TextMatrix(0, 7) = "Œ’„"
    .TextMatrix(0, 8) = "ﬂÊœ Œ’„"
    .TextMatrix(0, 9) = "«·‘Õ‰"
    .TextMatrix(0, 10) = "«·≈Ã„«·Ï"
    .TextMatrix(0, 11) = "ÿ—Ìﬁ… «·œ›⁄"
    
    .TextMatrix(0, 12) = "«·›—⁄"
    .TextMatrix(0, 13) = "«ŸÂ«— ··›—⁄"
    .TextMatrix(0, 14) = "SEND TIME "
    .TextMatrix(0, 15) = "„·«ÕŸ«  ÿ·»Ì…"
    .ColComboList(12) = StrListA("SELECT CODE , DESCA FROM FILE0_40 WHERE ONLINE = 1 ORDER BY CODE ", con_SERVER)
    .ColComboList(13) = "..."
    .ColComboList(19) = "..."
    .ColComboList(21) = "..."
    
    .TextMatrix(0, 16) = "»Ê‰ „»Ì⁄« "
    .TextMatrix(0, 17) = " «—ÌŒ"
    .TextMatrix(0, 18) = "»Ê·Ì’… ‘Õ‰"
    
    .TextMatrix(0, 20) = "«·€«¡ «·ÿ·»"
    .TextMatrix(0, 19) = "”»» «·≈·€«¡"
    .TextMatrix(0, 21) = " «—ÌŒ «” ·«„ «·„— Ã⁄"
    
    .ColDataType(1) = flexDTDate
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ColDataType(9) = flexDTDouble
    .ColDataType(10) = flexDTDouble
    
    .ColHidden(14) = True
    
    .ColWidth(0) = 800
    .ColWidth(1) = 1300
    .ColWidth(2) = 2000
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 4000
    .ColWidth(6) = 1000
    .ColWidth(7) = 900
    .ColWidth(8) = 800
    .ColWidth(9) = 800
    .ColWidth(10) = 900
    .ColWidth(11) = 1400
    .ColWidth(12) = 1800
    .ColWidth(13) = 1200
    .ColWidth(16) = 1000
    .ColWidth(15) = 2000
    .ColWidth(17) = 1300
    .ColWidth(18) = 2000
    .ColWidth(19) = 1300
    .ColWidth(20) = 1300
    .ColWidth(21) = 1300
    .ColDataType(21) = flexDTDate
    
    
    For nRow = 1 To .Rows - 1
'       .TextMatrix(nRow, 6) = Val(.TextMatrix(nRow, 6)) + Val(.TextMatrix(nRow, 7))
    Next nRow
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    '.ColDataType(0) = flexDTDate
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTCount, -1, 0, "#", vbYellow, , True, ""
    
    .Subtotal flexSTSum, -1, 6, "#", vbYellow, , True, ""
    .Subtotal flexSTSum, -1, 7, "#", vbYellow, , True, ""
    .Subtotal flexSTSum, -1, 9, "#", vbYellow, , True, ""
    .Subtotal flexSTSum, -1, 10, "#", vbYellow, , True, ""
    
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    closeCon con_SERVER
    If cBranch <> "00" Then closeCon con_MyShop
End Sub
Sub AddFromExel()
    Dim xl As New Excel.Application, nREcOrder As Double
    Dim xlsheet As Excel.Worksheet, nItem As String, lAddOrder As Boolean
    Dim xlwbook As Excel.Workbook, cItem As String
    Dim cFileName As String, cBarCode As String, cDoc_No As String
    Dim nCount As Double, cModel As String, nquant As Double, nPrice As Double, nPriceC As Double
    Dim cMail As String, nSubTotal As Double, nTotal As Double, nDisc As Double, nShip As Double, cDidc_Code As String, dDate As Date, cCity As String, cAddress1 As String, cAddress2 As String, cPhone As String, cSku As String, cName As String, cStreet As String, cPayment   As String, cShipping_City As String
    Dim aInsert As Variant, aInsert2 As Variant

    If MsgBox("≈÷«›… »Ì«‰«  „‰ „·› Excel", vbYesNo) = vbYes Then
        Common1.InitDir = ""
        Common1.FileName = ""
        Common1.Filter = "Excel (*.XLS*)|*.XLS*"
        Common1.ShowOpen
        cFileName = Common1.FileName
        If cFileName <> "" Then
            Set xlwbook = xl.Workbooks.Open(cFileName)
            Set xlsheet = xlwbook.Sheets.Item(1)
            With grid1
            I = 1
            Do While True
                I = I + 1
                Me.Caption = I
                If (xlsheet.Cells(I, 1)) <> "" Then
                    nCount = I
                Else
                    Exit Do
                End If
            Loop
            prog1.Visible = True
            prog1.Value = 0
            prog1.Max = nCount
            prog1.Min = 0
            .Rows = 1
            prog1.Min = 0
            prog1.Max = nCount
            For nRow = 2 To nCount
                prog1.Value = nRow
                If xlsheet.Cells(nRow, 1) <> "" Then
                    If cDoc_No <> Mid(xlsheet.Cells(nRow, 1), 2) Then
                        cDoc_No = Mid(xlsheet.Cells(nRow, 1), 2)
                        aInsert = AddFlag(Empty, "DOC_NO", addstring(cDoc_No))
                        For nCol = 2 To 90
                            Select Case xlsheet.Cells(1, nCol)
                            Case "Email"
                                    aInsert = AddFlag(aInsert, "E_MAIL", addstring(Trim(xlsheet.Cells(nRow, nCol))))
                                Case "Subtotal"
                                    aInsert = AddFlag(aInsert, "SUBTOTAL", Val(xlsheet.Cells(nRow, nCol)))
                                Case "Shipping"
                                    aInsert = AddFlag(aInsert, "SHIPPING", Val(xlsheet.Cells(nRow, nCol)))
                                Case "Total"
                                    aInsert = AddFlag(aInsert, "TOTAL", Val(xlsheet.Cells(nRow, nCol)))
                                Case "Discount Code"
                                    aInsert = AddFlag(aInsert, "DISCOUNT_CODE", addstring(Trim(xlsheet.Cells(nRow, nCol))))
                                Case "Discount Amount"
                                    aInsert = AddFlag(aInsert, "DISCOUNT", Val(xlsheet.Cells(nRow, nCol)))
                                Case "Created at"
                                    aInsert = AddFlag(aInsert, "DATE", addDate(DateValue(Mid(xlsheet.Cells(nRow, nCol), 1, 10))))
                                Case "Shipping Name"
                                    aInsert = AddFlag(aInsert, "NAME", addstring(Trim(xlsheet.Cells(nRow, nCol))))
                                Case "Shipping Street"
                                    cStreet = Trim(xlsheet.Cells(nRow, nCol))
                                    cStreet = Replace(cStreet, "'", " ")
                                    aInsert = AddFlag(aInsert, "STREET", addstring(cStreet))
                                Case "Shipping Address1"
                                    cAddress1 = Trim(xlsheet.Cells(nRow, nCol))
                                    cAddress1 = Replace(cAddress1, "'", " ")
                                    aInsert = AddFlag(aInsert, "Address1", addstring(cAddress1))
                                Case "Shipping Address2"
                                    cAddress2 = Trim(xlsheet.Cells(nRow, nCol))
                                    cAddress2 = Replace(cAddress1, "'", " ")
                                    aInsert = AddFlag(aInsert, "Address2", addstring(cAddress2))
                                Case "Shipping City"
                                    cCity = Trim(xlsheet.Cells(nRow, nCol))
                                    cCity = Replace(cCity, "'", " ")
                                    aInsert = AddFlag(aInsert, "CITY", addstring(cCity))
                                Case "Shipping Phone"
                                    cPhpne = Trim(xlsheet.Cells(nRow, nCol))
                                    cPhpne = Replace(cPhpne, " ", "")
                                    aInsert = AddFlag(aInsert, "PHONE", addstring(cPhpne))
                                Case "Payment Method"
                                    aInsert = AddFlag(aInsert, "Payment_Method", addstring(Trim(xlsheet.Cells(nRow, nCol))))
                                Case "Shipping Province Name"
                                    aInsert = AddFlag(aInsert, "Shipping_City", addstring(Trim(xlsheet.Cells(nRow, nCol))))
                                Case "Notes"
                                    aInsert = AddFlag(aInsert, "NOTES_order", addstring(Trim(xlsheet.Cells(nRow, nCol))))
                            End Select
                        Next nCol
                        lAddOrder = False
                        aRet = aGetDesca("SELECT DOC_NO , STORE FROM FILE6_90H WHERE DOC_NO = " & MyParn(cDoc_No), con_SERVER)
                        If UBound(aRet) = 0 Then
                            lAddOrder = True
                            con_SERVER.Execute addInsert(aInsert, "FILE6_90H")
                        Else
                            If TurnValue(aRet(2), Null, "") = "" Then
                                lAddOrder = True
                                con_SERVER.Execute addUpdate(aInsert, "FILE6_90H", " DOC_NO = " & MyParn(cDoc_No))
                            End If
                        End If
                        If lAddOrder Then
                            nREcOrder = nREcOrder + 1
                            con_SERVER.Execute " DELETE FROM FILE6_90 WHERE DOC_NO = " & MyParn(cDoc_No)
                        End If
                    End If
                    If lAddOrder Then
                        aInsert2 = AddFlag(Empty, "DOC_NO", addstring(cDoc_No))
                        cItem_NAME = ""
                        For nCol = 2 To 90
                            Select Case xlsheet.Cells(1, nCol)
                                Case "Lineitem price"
                                    aInsert2 = AddFlag(aInsert2, "PRICE", Val(xlsheet.Cells(nRow, nCol)))
                                
                                Case "Lineitem quantity"
                                    aInsert2 = AddFlag(aInsert2, "QUANT", Val(xlsheet.Cells(nRow, nCol)))
                                Case "Lineitem name"
                                    cItem_NAME = Trim(xlsheet.Cells(nRow, nCol))
                                    aInsert2 = AddFlag(aInsert2, "ITEM_NAME", addstring(Trim(xlsheet.Cells(nRow, nCol))))
                                Case "Lineitem pricename"
                                    aInsert2 = AddFlag(aInsert2, "PRICE", Val(xlsheet.Cells(nRow, nCol)))
                                Case "Lineitem sku"
                                    cSku = xlsheet.Cells(nRow, nCol)
                                    aInsert2 = AddFlag(aInsert2, "sku", addstring(Trim(xlsheet.Cells(nRow, nCol))))
                            End Select
                        Next nCol
                        
                        cItem = ""
                        Dim loctable As ADODB.Recordset
                        If Len(cSku) > 8 Then
                            Set loctable = ItemFind_BARCODE(cSku, con_SERVER)
                            If loctable.RecordCount > 0 Then cItem = loctable!Item
                        Else
                            cItem = cSku
                        End If
                        If cItem = "" Then MsgBox "ERROR SKU ORDER NO # " & cDoc_No & " - " & cItem_NAME
                        aInsert2 = AddFlag(aInsert2, "ITEM", addstring(cItem))
                        con_SERVER.Execute addInsert(aInsert2, "FILE6_90")
                    End If
                End If
            Next nRow
            MsgBox " „ «÷«›… «·ÿ·»Ì«  ⁄œœ " & nREcOrder
            End With
        End If
    End If
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If cBranch <> "00" Then Exit Sub
With grid1
    Select Case Col
        Case 2
            con_SERVER.Execute " UPDATE FILE6_90H SET [NAME] = " & addstring(.TextMatrix(Row, Col)) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
        Case 3
            con_SERVER.Execute " UPDATE FILE6_90H SET [PHONE] = " & addstring(.TextMatrix(Row, Col)) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
        Case 4
            con_SERVER.Execute " UPDATE FILE6_90H SET [CITY] = " & addstring(.TextMatrix(Row, Col)) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
        Case 5
            con_SERVER.Execute " UPDATE FILE6_90H SET [ADDRESS1] = " & addstring(.TextMatrix(Row, Col)) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
    End Select
End With
End Sub
Private Sub grid1_DblClick()
    MYLOAD2 grid1.TextMatrix(grid1.Row, 0)
End Sub
Private Sub MYLOAD2(pDoc_No)
'On Error GoTo myerror
Dim StoreTable As New ADODB.Recordset
Dim cBalNotShip As String
If cBranch = "00" Then
    If Check1.Value = 0 Then
        StoreTable.Open " SELECT * FROM FILE0_40 WHERE ONLINE = 1 ORDER BY CODE ", con_SERVER, adOpenStatic, adLockReadOnly, adCmdText
    Else
        StoreTable.Open " SELECT * FROM FILE0_40 WHERE ISONEST2 = 0 AND ISSTOP = 0 AND ISONEST = 0 AND BRANCH IS NOT NULL  AND CODE <> '000' ORDER BY CODE ", con_SERVER, adOpenStatic, adLockReadOnly, adCmdText
    End If
Else
    StoreTable.Open " SELECT * FROM FILE0_40 WHERE CODE = " & MyParn(cBranchStore), con_SERVER, adOpenStatic, adLockReadOnly, adCmdText
End If
Dim cString  As String, cStrStore As String
With StoreTable
    cStrStore = " , (SELECT SUM(BALITEM) FROM [BAL_ITEM_ORDERONLINE] WHERE [BAL_ITEM_ORDERONLINE].ITEM = FILE6_90.ITEM AND [BAL_ITEM_ORDERONLINE].STORE  = " & MyParn(StoreTable!CODE) & " ) AS 'Total All Order' "
    Do While Not .EOF
        cStrStore = cStrStore & " , (SELECT SUM([IN]-[OUT]) FROM FILE1_11 WHERE FILE1_11.ITEM = FILE6_90.ITEM AND FILE1_11.STORE  = " & MyParn(StoreTable!CODE) & " ) AS '" & StoreTable!DESCA & "'"
        .MoveNext
    Loop
End With

GRID2.Rows = 2
GRID2.FixedRows = 2

cBalNotShip = " ,(SELECT SUM(QUANT ) FROM Q_BAL_NOTSHIP WHERE Q_BAL_NOTSHIP.ITEM = FILE6_90.ITEM)"
With grid1
'                               0           1               2           3                                                                                   4               5               6            7                  8                   9           10          11
    cString = " SELECT    FILE6_90.ITEM, FILE6_90.SKU, FACT.DESCA , CASE WHEN FILE1_10.desca IS NULL THEN FILE6_90.ITEM_NAME ELSE FILE1_10.desca END , FILE1_10.COLOR, FILE1_10.SCAL, FILE6_90.QUANT, FILE6_90.PRICE, file6_90.ITEM_NAME, FILE6_90.ID " & cStrStore & cBalNotShip & " FROM            FILE6_90 LEFT  JOIN FILE1_10 ON FILE6_90.ITEM = FILE1_10.ITEM left  JOIN  FACT ON FILE1_10.code = FACT.CODE WHERE DOC_NO =  " & MyParn(pDoc_No)
    Set DATA2.Recordset = myRecordSet(cString, con_SERVER)
End With
FixGrid2
GRID2.Cell(flexcpAlignment, 0, 0, GRID2.Rows - 1, GRID2.Cols - 1) = 7
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Sub FixGrid2()
Dim LShowColAll As Boolean, LShowCol As Boolean
With GRID2
        
    .FixedRows = 2
    .RowHeight(0) = 1000
    .RowHidden(1) = True
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "SKU"
    .TextMatrix(0, 2) = "«·„’‰⁄"
    .TextMatrix(0, 3) = "«·’‰›"
    .TextMatrix(0, 4) = "«··Ê‰"
    .TextMatrix(0, 5) = "„ﬁ«”"
    .TextMatrix(0, 6) = "«·ﬂ„Ì…"
    .TextMatrix(0, 7) = "«·”⁄—"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
    .ColWidth(3) = 4000
    .ColWidth(4) = 1500
    .ColWidth(5) = 800
    .ColWidth(6) = 600
    .ColWidth(7) = 700
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    .ColWidth(.Cols - 1) = 1000
    .TextMatrix(0, .Cols - 1) = "—’Ìœ ‘Õ‰ €Ì— „” ·„"
    LShowColAll = False
    For nCol = 11 To .Cols - 1
        If Check1.Value <> 0 Then
            .TextMatrix(1, nCol) = IIf(GetBoolean("select online from file0_40 where desca = " & MyParn(.TextMatrix(0, nCol)), con_SERVER), 1, 0)
        Else
            .TextMatrix(1, nCol) = 1
        End If
        If cBranch = "00" Then
            .ColWidth(nCol) = 800
        Else
            .ColWidth(nCol) = 1200
        End If
        LShowCol = True
        For nRow = 2 To .Rows - 1
            If .ValueMatrix(nRow, nCol) = 0 Then
                LShowCol = False
            End If
        Next nRow
'        .ColHidden(nCol) = Not LShowCol Or .TextMatrix(1, nCol) = "0"
        If LShowCol And .TextMatrix(1, nCol) = "1" Then LShowColAll = True
    Next nCol
    If Not LShowColAll Then
        For nCol = 11 To .Cols - 1
            .ColHidden(nCol) = False
        Next nCol
    End If
    For nCol = 11 To .Cols - 1
        If Not .ColHidden(nCol) Then
            LShowCol = True
            For nRow = 2 To .Rows - 1
                If .ValueMatrix(nRow, nCol) = 0 Then
                    LShowCol = False
                End If
            Next nRow
            If LShowCol Then
                .Cell(flexcpBackColor, 0, nCol, .Rows - 1, nCol) = vbGreen
            Else
                .Cell(flexcpBackColor, 0, nCol, .Rows - 1, nCol) = &H8080FF
            End If
        End If
    Next nCol
    If cBranch <> "00" Then
        .Cols = .Cols + 1
        .ColWidth(.Cols - 1) = 1200
        .TextMatrix(0, .Cols - 1) = "—’Ìœ «·›—⁄"
        For nRow = 2 To .Rows - 1
            .TextMatrix(nRow, .Cols - 1) = LastBalance(GRID2.TextMatrix(nRow, 0), cBranchStore, con_MyShop)
        Next nRow
    End If
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .ColDataType(0) = flexDTDate
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTCount, -1, 0, "#", vbYellow, , True, ""
    .Subtotal flexSTSum, -1, 6, "#", vbYellow, , True, ""
    .Subtotal flexSTSum, -1, 8, "#", vbYellow, , True, ""
    .Select .Rows - 1, 0
    .ShowCell .Rows - 1, 0
    End With
End Sub

Private Sub grid1_EnterCell()
With grid1
    CMD_SEND.Visible = CMD_SEND.Visible = (cBranch <> "00") And lIsBoxOnline And cManBox <> "" And Not lSupperVisor And .TextMatrix(.Row, 16) = ""
    If (.Col = 2 Or .Col = 3 Or .Col = 4 Or .Col = 5 Or .Col = 12 Or .Col = 13 Or .Col = 19) And cBranch = "00" Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
    xNotes.Caption = .TextMatrix(.Row, 15)
    xUser.Caption = .TextMatrix(.Row, 13)
    xtime.Caption = .TextMatrix(.Row, 14)
End With
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Dim cNoteS As String
With grid1
    If Col = 13 And .TextMatrix(Row, 19) = "" Then
        If .TextMatrix(Row, 12) <> "" Then
            If MsgBox(" «ŸÂ«— ÿ·»Ì…  «Ê‰ ·«Ì‰ ··›—⁄ ", vbYesNo + vbDefaultButton2) = vbYes Then
                cNoteS = InputBox("", " „·«ÕŸ«  Œ«’… »«·ÿ·» ··›—⁄ ")
                
                .TextMatrix(Row, 13) = cusername
                .TextMatrix(Row, 14) = Date
                .TextMatrix(Row, 15) = cusername
                .TextMatrix(Row, 15) = cNoteS
                
                If Trim(.TextMatrix(Row, 12)) = "" Then
                    con_SERVER.Execute " UPDATE FILE6_90H SET STORE = NULL , SEND_USER = NULL , SEND_TIME = NULL , NOTES = NULL  WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
                Else
                    con_SERVER.Execute " UPDATE FILE6_90H SET STORE = " & addstring(.TextMatrix(Row, 12)) & " , SEND_USER = " & addstring(cusername) & " , SEND_TIME = GETDATE()  , NOTES= " & addstring(cNoteS) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
                End If
                xNotes.Caption = cNoteS
                xUser.Caption = cusername
                xtime.Caption = Date & " " & Time
            
            End If
        Else
            If MsgBox(" «·€«¡  —ÕÌ· «·ÿ·»Ì…  ··›—⁄ ", vbYesNo + vbDefaultButton2) = vbYes Then
                con_SERVER.Execute " UPDATE FILE6_90H SET STORE = NULL , SEND_USER = NULL , SEND_TIME = NULL , NOTES = NULL  WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
                xNotes.Caption = ""
                xUser.Caption = ""
                xtime.Caption = ""
            
                .TextMatrix(Row, 13) = ""
                .TextMatrix(Row, 14) = ""
                .TextMatrix(Row, 15) = ""
                .TextMatrix(Row, 15) = ""
            
            End If
        End If
    End If
    If Col = 19 And .TextMatrix(Row, 12) <> "" Then
        If MsgBox(" «·€«¡ «·ÿ·»Ì… „‰ «·⁄„Ì· ", vbYesNo + vbDefaultButton2) = vbYes Then
            cNoteS = InputBox("", " ”»» «·≈·€«¡ ")
            .TextMatrix(Row, 19) = cNoteS
            .TextMatrix(Row, 20) = Date
            con_SERVER.Execute " UPDATE FILE6_90H SET DelOrder_Date = " & DateSq(Date) & "  , NOTES = " & addstring(cNoteS) & " WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
        Else
            cNoteS = ""
            .TextMatrix(Row, 19) = ""
            .TextMatrix(Row, 20) = ""
            con_SERVER.Execute " UPDATE FILE6_90H SET DelOrder_Date = null , NOTES = null WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
        End If
    End If
    If Col = 21 And IsDate(.TextMatrix(Row, 20)) Then
        If MsgBox("  «—ÌŒ «” ·«„ «·„— Ã⁄ „‰ ‘—ﬂ… «·‘Õ‰ ", vbYesNo + vbDefaultButton2) = vbYes Then
            DDate2 = InputBox("", " «· «—ÌÕ ")
            .TextMatrix(Row, 21) = DDate2
            con_SERVER.Execute " UPDATE FILE6_90H SET DelOrder_Date2 = " & DateSq(sDate2) & "  WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
        Else
            con_SERVER.Execute " UPDATE FILE6_90H SET DelOrder_Date2 = null  WHERE DOC_NO = " & MyParn(.TextMatrix(Row, 0))
        End If
    End If
End With
End Sub
Private Function LoadConString_OnLine()
Exit Function
Dim cServerName As String, cUserId As String, cPassword As String, sCatalog_Online As String
'    cServerName = "junior-sql.database.windows.net"
'    cUserId = decrypt("062F8C407C77", "dr")
'    cPassword = decrypt("C77DA5F185954963DD1F902815", "dr")
'    sCatalog_Online = "JUNIOR_SHOP"
'    LoadConString_OnLine = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
'        & "catalog=" & sCatalog_Online & ";user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"
End Function
Private Function myreplace_Sales() As String
    Dim cDoc As String, cDoc2 As String, dSalesDate As String
    If lIsBranchStore Then
        dSalesDate = Format(GetDesca("SELECT DSALES FROM DSALES WHERE BRANCH = " & MyParn(cBranch), con_MyShop), "DD-MM-YYYY")
    Else
        dSalesDate = Format(GetDesca("SELECT DSALES FROM DSALES", con_MyShop), "DD-MM-YYYY")
    End If
    
    Dim aInsert(23, 1)
    Dim aInsert2(9, 1)
    
    aInsert(0, 0) = "Doc_No"
    aInsert(0, 1) = addstring(cDoc)

    aInsert(1, 0) = "code"
    aInsert(1, 1) = addstring("0000")

    aInsert(2, 0) = "[Date]"
    aInsert(2, 1) = addDate(dSalesDate)

    aInsert(3, 0) = "store"
    aInsert(3, 1) = addstring(cBranchStore)

    aInsert(4, 0) = "Discount"
    aInsert(4, 1) = Val(grid1.ValueMatrix(grid1.Row, 7))

    aInsert(5, 0) = "Box"
    aInsert(5, 1) = addstring(cManBox)


    aInsert(6, 0) = "userName"
    aInsert(6, 1) = addstring(cusername)

    aInsert(7, 0) = "branch"
    aInsert(7, 1) = addstring(cBranch)

    aInsert(8, 0) = "MAN"
    aInsert(8, 1) = addstring("0001")

    aInsert(9, 0) = "PHONE"
    aInsert(9, 1) = addstring(grid1.TextMatrix(grid1.Row, 3))

    aInsert(10, 0) = "isnew"
    aInsert(10, 1) = 1

    If Val(grid1.ValueMatrix(grid1.Row, 7)) <> 0 Then
        aInsert(11, 0) = "username_disc"
        aInsert(11, 1) = addstring("«Ê‰·«Ì‰")
        
        aInsert(12, 0) = "NOTE_DISC"
        aInsert(12, 1) = addstring(grid1.TextMatrix(grid1.Row, 8))
    
    Else
        aInsert(11, 0) = "username_disc"
        aInsert(11, 1) = addstring("")
        
        aInsert(12, 0) = "NOTE_DISC"
        aInsert(12, 1) = addstring("")
    End If


    aInsert(13, 0) = "ONLINE"
    aInsert(13, 1) = 2


    aInsert(14, 0) = "USER_IP"
    aInsert(14, 1) = addstring(GetComputerName)

    aInsert(15, 0) = "ADDRESS"
    aInsert(15, 1) = addstring(grid1.TextMatrix(grid1.Row, 4) & " " & grid1.TextMatrix(grid1.Row, 5))

    aInsert(16, 0) = "NAME"
    aInsert(16, 1) = addstring(grid1.TextMatrix(grid1.Row, 2))

    aInsert(17, 0) = "CHARGE2"
    aInsert(17, 1) = Val(grid1.ValueMatrix(grid1.Row, 9))

    aInsert(18, 0) = "ONLINE_DOC"
    aInsert(18, 1) = addstring(grid1.ValueMatrix(grid1.Row, 0))

    aRet = aGetDesca("SELECT CITY , Shipping_City , STREET FROM FILE6_90H WHERE DOC_NO = " & MyParn(grid1.TextMatrix(grid1.Row, 0)), con_SERVER)
          
    aInsert(19, 0) = "CITY"
    aInsert(19, 1) = addstring(aRet(1))
      
    aInsert(20, 0) = "Shipping_City"
    aInsert(20, 1) = addstring(aRet(2))
      
    aInsert(21, 0) = "STREET"
    aInsert(21, 1) = addstring(aRet(3))
      
    aInsert(22, 0) = "Payment_Method"
    aInsert(22, 1) = addstring(grid1.TextMatrix(grid1.Row, 11))
      
    aInsert(23, 0) = "SHIP"
    aInsert(23, 1) = addstring("09")
      

    cDoc = NewflagDoc(dSalesDate, cManBox, con_MyShop)
    aInsert(0, 0) = "Doc_No"
    aInsert(0, 1) = addstring(cDoc)
    
    con_MyShop.BeginTrans
    con_MyShop.Execute CreateInsert(aInsert, "FILE6_20H")
    
    For nRow = 2 To GRID2.Rows - 2
        aInsert2(0, 0) = "Doc_No"
        aInsert2(0, 1) = addstring(cDoc)
        
        aInsert2(1, 0) = "ITEM"
        aInsert2(1, 1) = addvalue(GRID2.TextMatrix(nRow, 0))
        
        aInsert2(2, 0) = "QUANT"
        aInsert2(2, 1) = Val(GRID2.ValueMatrix(nRow, 6))
        
        aInsert2(3, 0) = "PRICE"
        aInsert2(3, 1) = Val(GRID2.ValueMatrix(nRow, 7))

        aInsert2(4, 0) = "row"
        aInsert2(4, 1) = nRow

        aInsert2(5, 0) = "MAN"
        aInsert2(5, 1) = addstring("0001")

        aInsert2(6, 0) = "USER_IP"
        aInsert2(6, 1) = addstring(GetComputerNamecIpName)

        aInsert2(7, 0) = "Price_C2"
        aInsert2(7, 1) = Val(GRID2.ValueMatrix(nRow, 7))
        
        aInsert2(8, 0) = "DESCA2"
        aInsert2(8, 1) = addstring(GRID2.TextMatrix(nRow, 8))
        
        aInsert2(9, 0) = "SKU"
        aInsert2(9, 1) = addstring(GRID2.TextMatrix(nRow, 1))
        
        con_MyShop.Execute CreateInsert(aInsert2, "FILE6_20")
            
    Next nRow
    con_MyShop.CommitTrans
    con_MyShop.Execute " UPDATE FILE6_20 SET  Price_C = FILE1_10.PRICE , COST = FILE1_10.COSTITEM FROM FILE6_20 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE6_20.ITEM  WHERE DOC_NO = " & MyParn(cDoc)
    con_SERVER.Execute " UPDATE FILE6_90H SET SALES_DOC = " & addstring(cDoc) & " , SALES_DATE = " & addDate(dSalesDate) & " WHERE DOC_NO = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
    myreplace_Sales = cDoc
Exit Function
myerror:
MsgBox Err.Description
con_MyShop.RollbackTrans
Err.Clear
End Function

Private Sub grid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)

'    .TextMatrix(0, 0) = "ﬂÊœ"
'    .TextMatrix(0, 1) = "SKU"
'    .TextMatrix(0, 2) = "«·„’‰⁄"
'    .TextMatrix(0, 3) = "«·’‰›"
'    .TextMatrix(0, 4) = "«··Ê‰"
'    .TextMatrix(0, 5) = "„ﬁ«”"
'    .TextMatrix(0, 6) = "«·ﬂ„Ì…"
'    .TextMatrix(0, 7) = "«·”⁄—"

With GRID2
    If grid1.TextMatrix(grid1.Row, 16) = "" Then
        If .Col = 0 Then
            Dim loctable As ADODB.Recordset
            Set loctable = ItemFind(Val(.TextMatrix(Row, 0)), con_SERVER)
            If Not (loctable.EOF And loctable.BOF) Then
                GRID2.TextMatrix(Row, 1) = loctable!BARCODE13 & ""
                GRID2.TextMatrix(Row, 3) = loctable!DESCA & ""
                GRID2.TextMatrix(Row, 7) = loctable!price & ""
                            
                If Val(.TextMatrix(Row, 9)) <> 0 Then
                    con_SERVER.Execute " UPDATE FILE6_90 SET ITEM = " & addvalue(.TextMatrix(Row, 0)) & " , ITEM_NAME = " & addstring(.TextMatrix(Row, 3)) & " , PRICE = " & Val(.TextMatrix(Row, 7)) & " , SKU = " & addstring(.TextMatrix(Row, 1)) & " WHERE ID = " & Val(.TextMatrix(Row, 9))
                Else
                    GRID2.TextMatrix(Row, 6) = 1
                    cStr1 = "INSERT INTO FILE6_90 ( doc_no , sku , item , item_name , price , quant )" & _
                                "VALUES( " & _
                                addstring(grid1.TextMatrix(grid1.Row, 0)) & "," & _
                                addstring(.TextMatrix(.Row, 1)) & "," & _
                                addvalue(.TextMatrix(.Row, 0)) & "," & _
                                addstring(.TextMatrix(.Row, 3)) & "," & _
                                Val(.TextMatrix(.Row, 7)) & "," & _
                                Val(.TextMatrix(.Row, 6)) & _
                                ")"
                    con_SERVER.Execute cStr1
                End If
                MYLOAD2 (grid1.TextMatrix(grid1.Row, 0))
            Else
                MsgBox " »«—ﬂÊœ €Ì— ’ÕÌÕ "
                MYLOAD2 (grid1.TextMatrix(grid1.Row, 0))
            End If
        End If
        If .Col = 6 Then
            con_SERVER.Execute " UPDATE FILE6_90 SET QUANT = " & Val(.TextMatrix(Row, 6)) & " WHERE ID = " & Val(.TextMatrix(Row, 9))
        End If
        If .Col = 7 Then
            con_SERVER.Execute " UPDATE FILE6_90 SET PRICE = " & Val(.TextMatrix(Row, 7)) & " WHERE ID = " & Val(.TextMatrix(Row, 9))
        End If
    End If
End With
End Sub

Private Sub grid2_EnterCell()
With GRID2
    If (.Col = 0 Or .Col = 7 Or .Col = 6) And cBranch = "00" Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub Grid2_KeyUp(KeyCode As Integer, Shift As Integer)
With GRID2
    If .Col = 0 And cBranch = "00" Then
        If KeyCode = 112 Then ItemsLookupAll Me, oSearchItem
    End If
    If KeyCode = 45 And cBranch = "00" And Col = 0 Then
        GRID2.AddItem GRID2.Row
        GRID2.TextMatrix(GRID2.Row, 0) = ""
    End If
    If KeyCode = 46 And cBranch = "00" And grid1.TextMatrix(grid1.Row, 16) = "" Then
        If MsgBox("  Õ–›  «·’‰› „‰ «·ÿ·»Ì… ", vbYesNo + vbDefaultButton2) = vbYes Then
            con_SERVER.Execute " DELETE FROM FILE6_90 WHERE ID = " & Val(GRID2.TextMatrix(GRID2.Row, 9))
            .RemoveItem GRID2.Row
        End If
    End If
End With
End Sub
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = GRID2.Name Then
    GRID2.TextMatrix(GRID2.Row, 0) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    grid2_AfterEdit GRID2.Row, 0
    Unload oSearchItem
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub doprint_day()
Dim aHeader(1)
'f Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim SourchTable As New ADODB.Recordset, I As Double
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
cString = " SELECT        FILE6_90H.DOC_NO, FILE1_10.MODELFACT0 ,FILE6_90H.NAME,FILE6_90H.PHONE ,FILE1_10.MOSM,FILE6_90H.DATE, FILE6_90H.STREET, FILE6_90H.CITY, FILE6_90H.Shipping_City, FILE6_90H.Payment_Method, FILE6_90.ITEM, FACT.DESCA AS FACT_NAME , FILE1_10.[GROUP], FILE1_10.COLOR,  FILE1_10.SCAL, FILE6_90.PRICE, FILE6_90.QUANT , FILE1_10.DESCA , FILE6_90H.NOTES , FILE0_40.DESCA AS STORENAME  " & _
            " FROM          FILE6_90H INNER JOIN FILE6_90 ON FILE6_90H.DOC_NO = FILE6_90.DOC_NO INNER JOIN FILE1_10 ON FILE6_90.ITEM = FILE1_10.ITEM INNER JOIN FACT ON FILE1_10.code = FACT.CODE LEFT JOIN FILE0_40 ON FILE0_40.CODE = FILE6_90H.STORE WHERE  SALES_DOC IS NULL  "
    If xDoc_No.text <> "" Then cString = cString & " AND [DOC_NO] = " & MyParn(xDoc_No.text)
    If xPay.BoundText <> "" Then cString = cString & " AND [Payment_Method] = " & MyParn(xPay.text)
    If xStore.BoundText <> "" Then cString = cString & " AND [STORE] = " & MyParn(xStore.BoundText)
    If IsDate(xDate1.text) Then cString = cString & " AND [DATE] >= " & DateSq(xDate1.text)
    If IsDate(xDate2.text) Then cString = cString & " AND [DATE] <= " & DateSq(xDate2.text)
    If cBranch = "00" Then
        If xIssend(0).Value <> 0 Then cString = cString & " AND STORE IS NULL  and DelOrder_Date is null "
        If xIssend(1).Value <> 0 Then cString = cString & " AND STORE IS NOT  NULL AND DelOrder_Date IS NULL "
        If xIssend(3).Value <> 0 Then cString = cString & " AND DelOrder_Date IS NOT  NULL "
    Else
        cString = cString & " AND STORE = " & MyParn(cBranchStore)
    End If
    If xsales(0).Value <> 0 Then cString = cString & " AND SALES_DOC IS NULL "
    If xsales(1).Value <> 0 Then cString = cString & " AND SALES_DOC IS NOT NULL "
    
SourchTable.Open cString, con_SERVER, adOpenStatic, adLockReadOnly, adCmdText
nCount = SourchTable.RecordCount
prog1.Min = 0
prog1.Value = 0
If nCount > 0 Then
    SourchTable.MoveFirst
Else
    MsgBox " ·« ÌÊÃœ »Ì«‰«  "
    Exit Sub
End If
prog1.Max = nCount
With SourchTable
    Do While Not .EOF
        I = I + 1
        prog1.Value = I
        temptable.AddNew
        temptable!str4 = !doc_no
        temptable!Date1 = !Date
        temptable!STR7 = !Name
        temptable!str6 = !phone
        
        temptable!str14 = !Shipping_City
        temptable!str15 = !Street
        temptable!str16 = !Payment_Method
        
        
        temptable!str1 = !Item
        temptable!str11 = !MOSM
        temptable!str12 = !fact_name
        temptable!str13 = !modelfact0
        temptable!str3 = !DESCA
        temptable!str9 = !NOTES
        temptable!str2 = !STORENAME
        
        temptable!str8 = !color
        temptable!str5 = !scal
        
        temptable!VAL1 = !Quant
        temptable!val3 = !price
        If cBranch <> "00" Then temptable!VAL4 = LastBalance(!Item, cBranchStore, con_MyShop)
        temptable!STR19 = " ÿ·»«  «Ê‰ ·«Ì‰  €Ì— „‰›–… " & xStore.text
        
        temptable.Update
    .MoveNext
Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\rep_order_online.rpt"
Main.REPORT1.DataFiles(0) = tempFile
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub

