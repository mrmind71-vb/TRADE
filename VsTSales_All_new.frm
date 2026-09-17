VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form VSTSALES_ALL 
   Appearance      =   0  'Flat
   ClientHeight    =   10365
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
   ScaleHeight     =   10365
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame fmBranch 
      Height          =   645
      Left            =   13905
      RightToLeft     =   -1  'True
      TabIndex        =   77
      Top             =   2880
      Width           =   2445
      Begin VB.CheckBox chkStoped 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "≈ŸÂ«— ›—Ê⁄ „ Êﬁ›…"
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
         Height          =   240
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   225
         Width           =   2085
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   2085
      Left            =   -5355
      RightToLeft     =   -1  'True
      TabIndex        =   73
      Top             =   585
      Visible         =   0   'False
      Width           =   5415
      Begin MSDataListLib.DataCombo XDOC_DISC 
         Height          =   315
         Left            =   90
         TabIndex        =   74
         Top             =   810
         Visible         =   0   'False
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label xModelNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label6"
         Height          =   690
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   225
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label2 
         Caption         =   "Œ’„ ⁄—Ê÷ :"
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
         Index           =   10
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   900
         Visible         =   0   'False
         Width           =   1230
      End
   End
   Begin VB.Frame Frame7 
      Height          =   645
      Left            =   16380
      RightToLeft     =   -1  'True
      TabIndex        =   72
      Top             =   2880
      Width           =   3795
      Begin VB.OptionButton xonest 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.OptionButton xonest 
         Alignment       =   1  'Right Justify
         Caption         =   "«„«‰« "
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
         Left            =   1215
         RightToLeft     =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   225
         Width           =   825
      End
      Begin VB.OptionButton xonest 
         Alignment       =   1  'Right Justify
         Caption         =   "€Ì— «„«‰« "
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   225
         Width           =   1185
      End
   End
   Begin VB.TextBox XTEXT1 
      Alignment       =   1  'Right Justify
      Height          =   2175
      Left            =   5085
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Text            =   "VsTSales_All_new.frx":0000
      Top             =   135
      Width           =   3840
   End
   Begin VB.Frame Frame6 
      Height          =   600
      Left            =   5085
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   2925
      Width           =   3840
      Begin VB.OptionButton SEC 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„Ê”„"
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
         Index           =   0
         Left            =   2835
         RightToLeft     =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   180
         Width           =   870
      End
      Begin VB.OptionButton SEC 
         Alignment       =   1  'Right Justify
         Caption         =   "«·‰Ê⁄"
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
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   180
         Width           =   735
      End
      Begin VB.OptionButton SEC 
         Alignment       =   1  'Right Justify
         Caption         =   "›∆… ⁄„—Ì…"
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
         Index           =   2
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
      End
      Begin VB.OptionButton SEC 
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
         Height          =   270
         Index           =   3
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   180
         Value           =   -1  'True
         Width           =   645
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   465
      Left            =   5085
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   1620
      Visible         =   0   'False
      Width           =   3840
      Begin MSAdodcLib.Adodc data25 
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
      Begin MSAdodcLib.Adodc data26 
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
      Begin MSAdodcLib.Adodc data27 
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
      Begin MSAdodcLib.Adodc data28 
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
      Begin MSAdodcLib.Adodc data29 
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
      Begin MSAdodcLib.Adodc data30 
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
      Begin MSAdodcLib.Adodc data31 
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
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6450
      Left            =   45
      TabIndex        =   51
      Top             =   3555
      Width           =   20175
      _ExtentX        =   35586
      _ExtentY        =   11377
      _Version        =   393216
      Tabs            =   12
      Tab             =   10
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "›—Ê⁄"
      TabPicture(0)   =   "VsTSales_All_new.frx":0006
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "GRID0"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "‘Â—Ï"
      TabPicture(1)   =   "VsTSales_All_new.frx":0022
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grid1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "ÌÊ„Ï"
      TabPicture(2)   =   "VsTSales_All_new.frx":003E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GRID2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "›∆«  ⁄„—Ì…"
      TabPicture(3)   =   "VsTSales_All_new.frx":005A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grid3"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "„Ã„Ê⁄« "
      TabPicture(4)   =   "VsTSales_All_new.frx":0076
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "grid4"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "«›”«„"
      TabPicture(5)   =   "VsTSales_All_new.frx":0092
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "grid5"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "„’«‰⁄"
      TabPicture(6)   =   "VsTSales_All_new.frx":00AE
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "grid6"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "„Ê—œÌ‰"
      TabPicture(7)   =   "VsTSales_All_new.frx":00CA
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "grid7"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "„Õ„Ê⁄«  „ﬁ«”« "
      TabPicture(8)   =   "VsTSales_All_new.frx":00E6
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "grid8"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "„ﬁ«”« "
      TabPicture(9)   =   "VsTSales_All_new.frx":0102
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "grid9"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "„ÊœÌ·« "
      TabPicture(10)  =   "VsTSales_All_new.frx":011E
      Tab(10).ControlEnabled=   -1  'True
      Tab(10).Control(0)=   "grid10"
      Tab(10).Control(0).Enabled=   0   'False
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "ÌÊ„Ï ›—Ê⁄"
      TabPicture(11)  =   "VsTSales_All_new.frx":013A
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "GRID11"
      Tab(11).Control(0).Enabled=   0   'False
      Tab(11).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid10 
         Height          =   5685
         Left            =   90
         TabIndex        =   52
         Top             =   675
         Width           =   19860
         _cx             =   35031
         _cy             =   10028
         _ConvInfo       =   1
         Appearance      =   1
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
      Begin VSFlex7Ctl.VSFlexGrid grid8 
         Height          =   5685
         Left            =   -74910
         TabIndex        =   53
         Top             =   675
         Width           =   19860
         _cx             =   35031
         _cy             =   10028
         _ConvInfo       =   1
         Appearance      =   1
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
         Rows            =   2
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
      Begin VSFlex7Ctl.VSFlexGrid grid7 
         Height          =   5685
         Left            =   -74910
         TabIndex        =   56
         Top             =   675
         Width           =   19815
         _cx             =   34951
         _cy             =   10028
         _ConvInfo       =   1
         Appearance      =   1
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
      Begin VSFlex7Ctl.VSFlexGrid grid6 
         Height          =   5685
         Left            =   -74910
         TabIndex        =   57
         Top             =   675
         Width           =   19860
         _cx             =   35031
         _cy             =   10028
         _ConvInfo       =   1
         Appearance      =   1
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
      Begin VSFlex7Ctl.VSFlexGrid grid5 
         Height          =   5685
         Left            =   -74910
         TabIndex        =   58
         Top             =   675
         Width           =   19860
         _cx             =   35031
         _cy             =   10028
         _ConvInfo       =   1
         Appearance      =   1
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
         Cols            =   7
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
      Begin VSFlex7Ctl.VSFlexGrid grid4 
         Height          =   5685
         Left            =   -74910
         TabIndex        =   59
         Top             =   675
         Width           =   19860
         _cx             =   35031
         _cy             =   10028
         _ConvInfo       =   1
         Appearance      =   1
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
         Cols            =   7
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
      Begin VSFlex7Ctl.VSFlexGrid grid3 
         Height          =   5685
         Left            =   -74910
         TabIndex        =   60
         Top             =   675
         Width           =   19860
         _cx             =   35031
         _cy             =   10028
         _ConvInfo       =   1
         Appearance      =   1
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
         Cols            =   7
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
      Begin VSFlex7Ctl.VSFlexGrid GRID2 
         Height          =   5685
         Left            =   -74910
         TabIndex        =   61
         Top             =   675
         Width           =   19860
         _cx             =   35031
         _cy             =   10028
         _ConvInfo       =   1
         Appearance      =   1
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
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   5730
         Left            =   -74910
         TabIndex        =   62
         Top             =   675
         Width           =   19905
         _cx             =   35110
         _cy             =   10107
         _ConvInfo       =   1
         Appearance      =   1
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
      Begin VSFlex7Ctl.VSFlexGrid GRID0 
         Height          =   5685
         Left            =   -74910
         TabIndex        =   63
         Top             =   675
         Width           =   19860
         _cx             =   35031
         _cy             =   10028
         _ConvInfo       =   1
         Appearance      =   1
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
         Cols            =   14
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
      Begin VSFlex7Ctl.VSFlexGrid GRID11 
         Height          =   5685
         Left            =   -74910
         TabIndex        =   66
         Top             =   675
         Width           =   19860
         _cx             =   35031
         _cy             =   10028
         _ConvInfo       =   1
         Appearance      =   1
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
      Begin VSFlex7Ctl.VSFlexGrid grid9 
         Height          =   5640
         Left            =   -74910
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   675
         Width           =   19815
         _cx             =   34951
         _cy             =   9948
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   11
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
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   46
      Top             =   9840
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame Frame2 
      Height          =   555
      Left            =   5085
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   2385
      Width           =   3840
      Begin VB.OptionButton xall 
         Alignment       =   1  'Right Justify
         Caption         =   "«·ﬂ·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   555
      End
      Begin VB.OptionButton xall 
         Alignment       =   1  'Right Justify
         Caption         =   " ÊﬂÌ·«  ›ﬁÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   225
         Width           =   1095
      End
      Begin VB.OptionButton xall 
         Alignment       =   1  'Right Justify
         Caption         =   "›—Ê⁄ ›ﬁÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   225
         Width           =   960
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
      Height          =   150
      Left            =   -1305
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   3555
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Frame Frame1 
      Height          =   2835
      Left            =   8955
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   45
      Width           =   11235
      Begin VB.CheckBox XISONLINE 
         Alignment       =   1  'Right Justify
         Caption         =   "ONLINE"
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
         Height          =   240
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1350
         Width           =   915
      End
      Begin VB.CheckBox XOutlet 
         Alignment       =   1  'Right Justify
         Caption         =   "Outlet "
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
         Height          =   240
         Left            =   1215
         RightToLeft     =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1350
         Width           =   870
      End
      Begin VB.TextBox XMOSM2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   135
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   945
         Width           =   735
      End
      Begin VB.ComboBox xmosm22 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   540
         Width           =   915
      End
      Begin VB.TextBox XCOUNT 
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
         Left            =   315
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   2070
         Width           =   1815
      End
      Begin VB.TextBox xdoc_no 
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
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   1410
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "»œÊ‰ ≈Ã„«·Ï „Ê—œÌ‰"
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
         Left            =   6750
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2430
         Value           =   1  'Checked
         Width           =   2580
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
         Left            =   5985
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   1815
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
         Left            =   7845
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1590
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   5985
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
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   5985
         TabIndex        =   3
         Top             =   945
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
         Left            =   2250
         TabIndex        =   9
         Top             =   573
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   315
         Left            =   5985
         TabIndex        =   2
         Top             =   585
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   2250
         TabIndex        =   38
         Top             =   1329
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   2250
         TabIndex        =   10
         Top             =   951
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xcust 
         Height          =   315
         Left            =   6000
         TabIndex        =   5
         Top             =   1667
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGrCust 
         Height          =   315
         Left            =   5985
         TabIndex        =   6
         Top             =   2025
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xage 
         Height          =   315
         Left            =   2835
         TabIndex        =   11
         Top             =   1710
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xsex 
         Height          =   315
         Left            =   315
         TabIndex        =   17
         Top             =   1710
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XSHIP 
         Height          =   315
         Left            =   90
         TabIndex        =   12
         Top             =   180
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "‘—ﬂ… «·‘Õ‰ "
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
         Index           =   11
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   225
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "‰Ê⁄ «·„Ê”„ :"
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
         Index           =   9
         Left            =   1035
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1035
         Width           =   1050
      End
      Begin VB.Label Label2 
         Caption         =   "„Ê”„ œ«Œ·Ï"
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
         Index           =   8
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   585
         Width           =   1140
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
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   1755
         Width           =   555
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
         Index           =   4
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   1755
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«⁄·Ï ﬁÌ„… „»Ì⁄«  "
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   2115
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "›« Ê—… —ﬁ„ :"
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
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "„Ã„Ê⁄… „Ê—œÌ‰:"
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
         Index           =   2
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "«·⁄„Ì· :"
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
         TabIndex        =   41
         Top             =   1665
         Width           =   1140
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
         TabIndex        =   40
         Top             =   1020
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "«·›—⁄ :"
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
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1395
         Width           =   1230
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
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   270
         Width           =   765
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
         Index           =   6
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   630
         Width           =   1230
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
         TabIndex        =   34
         Top             =   974
         Width           =   1545
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
         Index           =   5
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1356
         Width           =   1410
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
         TabIndex        =   32
         Top             =   585
         Width           =   1545
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   30
      Top             =   10035
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data4 
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
   Begin MSAdodcLib.Adodc DATA21 
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
   Begin MSAdodcLib.Adodc DATA23 
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
   Begin MSAdodcLib.Adodc data8 
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
   Begin MSAdodcLib.Adodc data9 
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
   Begin MSAdodcLib.Adodc data10 
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
   Begin MSAdodcLib.Adodc DATA24 
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
   Begin MSAdodcLib.Adodc DATA11 
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
   Begin MSAdodcLib.Adodc DATA12 
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
   Begin MSAdodcLib.Adodc DATA13 
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
   Begin MSAdodcLib.Adodc DATA22 
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
   Begin MSAdodcLib.Adodc DATA14 
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
   Begin MSAdodcLib.Adodc data33 
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
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   8955
      RightToLeft     =   -1  'True
      TabIndex        =   67
      Top             =   2835
      Width           =   4920
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1230
         Picture         =   "VsTSales_All_new.frx":0156
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3600
         Picture         =   "VsTSales_All_new.frx":2941
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "VsTSales_All_new.frx":4E33
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2430
         Picture         =   "VsTSales_All_new.frx":729F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   135
         Width           =   1185
      End
   End
End
Attribute VB_Name = "VSTSALES_ALL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearch As New Search3, Search31 As New Search3
Dim con As New ADODB.Connection
Private Sub cmd_excel_Click()
End Sub
Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    Dim cHead3 As String
    cHead1 = "≈Ã„«·Ï „»Ì⁄«  „ÊœÌ·«  ·› —… " & XSTORE.text
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.text, "DD-MM-YYYY") & " Õ Ï  «—ÌŒ " & Format(xdate2.text, "DD-MM-YYYY")
    cHead3 = ""
    If xFact.BoundText <> "" Then cHead3 = cHead3 & " „’‰⁄ " & xFact.text
    If xSupp.BoundText <> "" Then cHead3 = cHead3 & " „Ê—œ " & xSupp.text & "  " & xGrCust.text
    If xGroup.BoundText <> "" Then cHead3 = cHead3 & " „Ã„Ê⁄…" & xGroup.text & "  " & xage.text & "  " & xsex.text
    If xSection.text <> "" Then cHead3 = cHead3 & " ﬁ”„ " & xSection.text
    
    If xMosm.BoundText <> "" Then cHead3 = cHead3 & " „Ê”„ " & xMosm.text
    
    PrintGrdNew.DOPRINT grid1, 1, , cHead1, cHead2, cHead3, , False, 8
    PrintGrdNew.Show 1
End Sub
Private Sub cmdExel_Click()
If SSTab1.Tab = 0 Then ToFileExel2 GRID0, , , , , 1.1, , , , , , Me
If SSTab1.Tab = 1 Then ToFileExel2 grid1, , , , , 1.1, , , , , , Me
If SSTab1.Tab = 2 Then ToFileExel2 GRID2, , , , , 1.1, , , , , , Me
If SSTab1.Tab = 3 Then ToFileExel2 grid3, , , , , 1.1, , , , , , Me
If SSTab1.Tab = 4 Then ToFileExel2 grid4, , , , , 1.1, , , , , , Me
If SSTab1.Tab = 5 Then ToFileExel2 grid5, , , , , 1.1, , , , , , Me
If SSTab1.Tab = 6 Then ToFileExel2 grid6, , , , , 1.1, , , , , , Me
If SSTab1.Tab = 7 Then ToFileExel2 grid7, , , , , 1.1, , , , , , Me
If SSTab1.Tab = 8 Then ToFileExel2 grid8, , , , , 1.1, , , , , , Me
If SSTab1.Tab = 9 Then ToFileExel2 grid9, , , , , 1.1, , , , , , Me
If SSTab1.Tab = 10 Then ToFileExel2 grid10, , , , , 1.1, , , , , , Me
If SSTab1.Tab = 11 Then ToFileExel2 GRID11, , , , , 1.1, , , , , , Me
End Sub

Private Sub cmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
'    On Error GoTo myerror
    If SSTab1.Tab = 0 Then
        myload0
    ElseIf SSTab1.Tab = 1 Then
        MYLOAD1
    ElseIf SSTab1.Tab = 2 Then
        MYLOAD2
    ElseIf SSTab1.Tab = 3 Then
        MYLOAD3
    ElseIf SSTab1.Tab = 4 Then
        MYLOAD4
    ElseIf SSTab1.Tab = 5 Then
        myload5
    ElseIf SSTab1.Tab = 6 Then
        myload6
    ElseIf SSTab1.Tab = 7 Then
        myload7
    ElseIf SSTab1.Tab = 8 Then
        myload8
    ElseIf SSTab1.Tab = 9 Then
        myloadGrd9
    ElseIf SSTab1.Tab = 10 Then
        myload10
    ElseIf SSTab1.Tab = 11 Then
        MYLOAD11
    End If
    Exit Sub
myerror:
    MsgBox Err.Description
End Sub
Private Sub Form_Load()
    openCon con
    
    LoadText Me
    
    
    XTEXT1.Visible = bSupermode Or lSupperVisor
        
    'FillCheckbox chkType, "SELECT CODE,DESCA FROM BRANCH_TYPE ORDER BY CODE", con
        
    xmosm22.AddItem " "
    xmosm22.AddItem "S"
    xmosm22.AddItem "W"
    xmosm22.AddItem "M"
    
    Set DATA11.Recordset = myRecordSet("Select Code,DescA From File0_82  order by Desca", con)
    Set xage.RowSource = DATA11
    xage.ListField = "Desca"
    xage.BoundColumn = "Code"
    
    Set DATA12.Recordset = myRecordSet("Select Code,DescA From File0_81  order by Desca", con)
    Set xsex.RowSource = DATA12
    xsex.ListField = "Desca"
    xsex.BoundColumn = "Code"
    
    Set data9.Recordset = myRecordSet("Select Code,DescA From File3_10  order by Desca", con)
    Set xcust.RowSource = data9
    xcust.ListField = "Desca"
    xcust.BoundColumn = "Code"
    
    Set data1.Recordset = myRecordSet("SELECT * FROM FILE4_50 ", con)
    Set xGrCust.RowSource = data1
    xGrCust.ListField = "Desca"
    xGrCust.BoundColumn = "Code"
    
    Set data8.Recordset = myRecordSet("Select Code,DescA From File1_10SC order by Desca", con)
    Set xSection.RowSource = data8
    xSection.ListField = "Desca"
    xSection.BoundColumn = "Code"
    
    
     Set DATA7.Recordset = myRecordSet("Select Code,DescA From STORE_BR WHERE " & cUserStore, con)
    
    Set XSTORE.RowSource = DATA7
    XSTORE.ListField = "Desca"
    XSTORE.BoundColumn = "Code"
    If cBranch <> "00" Then
        XSTORE.BoundText = cBranchStore
        XSTORE.Enabled = False
    End If
    
    
    Set DATA2.Recordset = myRecordSet("Select Code,DescA From File4_10 order by Desca", con)
    Set xSupp.RowSource = DATA2
    xSupp.ListField = "Desca"
    xSupp.BoundColumn = "Code"
    
    Set data3.Recordset = myRecordSet("Select Code,DescA From File1_50 ORDER BY DESCA", con)
    Set xGroup.RowSource = data3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    Set data4.Recordset = myRecordSet("Select mosm ,descA From mosm ORDER BY date DESC ", con)
    Set xMosm.RowSource = data4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    
    
    Set data5.Recordset = myRecordSet("Select code ,desca From fact ORDER BY DESCA ", con)
    Set xFact.RowSource = data5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    
    Set DATA13.Recordset = myRecordSet("Select DOC_NO ,DescA From File0_90H order by DOC_NO DESC ", con)
    Set XDOC_DISC.RowSource = DATA13
    XDOC_DISC.ListField = "Desca"
    XDOC_DISC.BoundColumn = "DOC_NO"
    
    Set DATA14.Recordset = myRecordSet("SELECT * FROM SHIP ", con)
    Set XSHIP.RowSource = DATA14
    XSHIP.ListField = "Desca"
    XSHIP.BoundColumn = "Code"
    
    
    Set GRID0.DataSource = data20
    Set grid1.DataSource = DATA21
    Set GRID2.DataSource = DATA22
    Set grid3.DataSource = DATA23
    Set grid4.DataSource = DATA24
    Set grid5.DataSource = data25
    Set grid6.DataSource = data26
    Set grid7.DataSource = data27
    Set grid8.DataSource = data28
    Set grid9.DataSource = data29
    Set grid10.DataSource = data30
    Set GRID11.DataSource = data31

    FIXGRID0
    FixGrid1
    FixGrid2
    FixGrid3
    FIXGRID4
    FIXGRID5
    
    FIXGRID6
    FIXGRID7
    FIXGRID8
    FIXGRID9
    FIXGRID10
    FixGrid11

    fmBranch.Visible = lSupperVisor
End Sub
Private Sub myload10()
Dim i As Double
Dim cString  As String, cWhere As String
With grid10
'                           0               1                 2                3               4                5
    
    If Val(XCOUNT.text) > 0 Then
        cString = " SELECT TOP " & Val(XCOUNT.text) & "  "
    Else
        cString = " SELECT "
    End If
        
    cString = cString & " ' ' AS N1 ," & _
             "FACT.desca," & _
             " file1_10.ModelNo," & _
             "file1_10.mosm," & _
             "file1_10.MODELFACT0," & _
             "file1_10.desca ,   " & _
            " SUM(VW_FILE6_20.QUANT) AS t_q," & _
            " SUM(VW_FILE6_20.TOTAL_NET) AS t_sales," & _
            " SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
            " SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft ," & _
            " CASE WHEN SUM(vw_FILE6_20.TOTAL_NET) <> 0  THEN ROUND(SUM(vw_FILE6_20.ITEM_PROFIT)  / SUM(vw_FILE6_20.TOTAL_NET),4) ELSE 0 END * 100 AS RAT_PROFIT," & _
            " SUM(vw_FILE6_20.ITEM_DISCOUNT_C)," & _
            " SUM(vw_FILE6_20.ITEM_DISCOUNT)," & _
            " SUM(vw_FILE6_20.TOTAL_C)," & _
            " CASE WHEN SUM(vw_FILE6_20.TOTAL_C) <> 0 THEN ROUND((SUM(vw_FILE6_20.TOTAL_C) - SUM(vw_FILE6_20.ITEM_COST)) / SUM(vw_FILE6_20.TOTAL_C),4) ELSE 0 END  * 100 AS CUST_RATE" & _
            " FROM FILE1_10 INNER JOIN VW_FILE6_20 ON FILE1_10.ITEM = VW_FILE6_20.ITEM" & _
            " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
            " inner join file4_10 on file4_10.code = file1_10.code" & _
            " INNER JOIN STORE_BR ON VW_FILE6_20.STORE = STORE_BR.CODE"
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    
    If xcust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.CUST = " & MyParn(xcust.BoundText)
    If xMosm.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    
    If xmosm22.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
    If xsex.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xage.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    
    If XDOC_DISC.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.DOC_DISC = " & MyParn(XDOC_DISC.BoundText)
    
    If XOutlet.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " file1_10.Outlet = 1 "
    
    If xSupp.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGrCust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
    If xGroup.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    
    
    If xdoc_no.text <> "" Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.doc_no = " & MyParn(xdoc_no.text)
    If XSTORE.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " STORE = " & MyParn(XSTORE.BoundText)
    If xSection.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[Section] = " & MyParn(xSection.BoundText)
    If IsDate(xDate1.text) Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.[DATE] <= " & DateSq(xdate2.text)
    
    If xall(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.ISBRANCH = 0 "
    
    If xonest(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 0 "
    If xonest(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 1 "
    If XISONLINE.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.online <> 0 "
    If XSHIP.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.[SHIP] = " & MyParn(XSHIP.BoundText)
     
    If cWhere <> "" Then cString = cString & " WHERE " & cWhere

    cString = cString & " GROUP BY file1_10.ModelNo, FACT.desca , file1_10.mosm, file1_10.MODELFACT0, file1_10.desca" & _
                        " ORDER BY t_sales DESC "
    Set data30.Recordset = mycmd(cString, con)
    XTEXT1.text = cString

End With
FIXGRID10
'If grid1.Rows > 0 Then ' grid1.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FIXGRID10()
With grid10
    .RowHeight(0) = 600
    .WordWrap = True
    .Cols = 15
    .TextMatrix(0, 0) = "„"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .TextMatrix(0, 2) = "„ÊœÌ·"
    .TextMatrix(0, 3) = "„Ê”„"
    .TextMatrix(0, 4) = "„ÊœÌ· "
    .TextMatrix(0, 5) = "«·’‰›"
    
    .TextMatrix(0, 6) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 7) = "ﬁÌ„… „»Ì⁄« "
    
    .TextMatrix(0, 8) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 9) = "—»Õ „»Ì⁄« "
    .TextMatrix(0, 10) = "‰”»… «·—»Õ"
    .TextMatrix(0, 11) = "Œ’„ «Êﬂ«“ÊÌ‰"
    .TextMatrix(0, 12) = "Œ’„ »Ê‰"
    
    .TextMatrix(0, 13) = "ﬁÌ„… »”⁄— „” Â·ﬂ"
    .TextMatrix(0, 14) = "‰”»… „” Â·ﬂ"
    
    .ColWidth(0) = 2500
    
    .ColWidth(0) = 500
    .ColWidth(2) = 0
    .ColWidth(3) = 800
    .ColWidth(4) = 1500
    .ColWidth(5) = 2800
    
    .ColWidth(6) = 1300
    .ColWidth(7) = 1300
    .ColWidth(8) = 1300
    .ColWidth(9) = 1300
    .ColWidth(10) = 1300
    .ColWidth(11) = 1300
    .ColWidth(12) = 1300
    .ColWidth(13) = 1300
    .ColWidth(14) = 800
    
    .ColHidden(8) = Not bOpt10
    .ColHidden(9) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
   
    For nRow = 1 To .Rows - 1
        .TextMatrix(nRow, 0) = nRow
        nRate = 0
        '.TextMatrix(nRow, 9) = Round(Val(.TextMatrix(nRow, 7)) - Val(.TextMatrix(nRow, 8)), 2)
        
'        If Val(.TextMatrix(nRow, 8)) <> 0 Then nRate = Round(Val(.TextMatrix(nRow, 9)) / Val(.TextMatrix(nRow, 8)) * 100, 2)
'        .TextMatrix(nRow, 10) = nRate
'        If Val(.TextMatrix(nRow, 13)) <> 0 Then
'            NP = Val(.TextMatrix(nRow, 13)) - Val(.TextMatrix(nRow, 8))
'            nRate = Round((NP / Val(.TextMatrix(nRow, 13))) * 100, 2)
'            .TextMatrix(nRow, 14) = nRate
'        End If
    Next nRow
    For nCol = 6 To .Cols - 2
        .ColFormat(nCol) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol) = Format(.TextMatrix(nRow, nCol), "#0.00")
        Next nRow
    Next nCol
    
    .SubtotalPosition = flexSTAbove
    For i = 6 To .Cols - 2
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
'        If Check1.Value = 0 Then
'            .Subtotal flexSTSum, 0, i, "#0", vbRed, vbYellow, True, "  "
'        End If
    Next i
    If .Rows > 1 Then
        If Val(.TextMatrix(1, 7)) <> 0 Then .TextMatrix(1, 10) = Round(Val(.TextMatrix(1, 9)) / Val(.TextMatrix(1, 7)) * 100, 2)
        If .ValueMatrix(1, 13) <> 0 Then .TextMatrix(1, 14) = Round(((.ValueMatrix(1, 13) - .ValueMatrix(1, 8)) / .ValueMatrix(1, 13)) * 100, 2)
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
    End With
End Sub
Private Sub Form_Resize()
Dim nHeight As Long
nHeight = Me.Height - SSTab1.Top - 1000
If nHeight < 5000 Then Exit Sub
SSTab1.Height = Me.Height - SSTab1.Top - 1000
GRID0.Height = SSTab1.Height - 750
grid1.Height = SSTab1.Height - 750
GRID2.Height = SSTab1.Height - 750
grid3.Height = SSTab1.Height - 750
grid4.Height = SSTab1.Height - 750
grid5.Height = SSTab1.Height - 750
grid6.Height = SSTab1.Height - 750
grid7.Height = SSTab1.Height - 750
grid8.Height = SSTab1.Height - 750
grid9.Height = SSTab1.Height - 750
grid10.Height = SSTab1.Height - 750
GRID11.Height = SSTab1.Height - 750
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SaveText Me
    closeCon con
End Sub

Private Sub XDOC_DISC_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        DocDisc_Look
    End If
End Sub

Private Sub xSection_LostFocus()
'    DATA3.ConnectionString = strCon
'    If XSECTION.BoundText = "" Then
'        DATA3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
'    Else
'        DATA3.RecordSource = "Select Code,DescA From File1_50 where [group] = " & Val(XSECTION.BoundText) & " ORDER BY DESCA"
'    End If
'    Set xGroup.RowSource = DATA3
'    xGroup.ListField = "Desca"
'    xGroup.BoundColumn = "Code"
'    DATA3.Refresh
End Sub

Private Sub xSupp_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then SuppLookupAll Me, oSearch
End Sub
Sub myProc()
If ActiveControl.Name = xSupp.Name Then
    xSupp.BoundText = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    oSearch.Hide
End If
If ActiveControl.Name = xFact.Name Then
    xFact.BoundText = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End If
If ActiveControl.Name = XDOC_DISC.Name Then
    XDOC_DISC.BoundText = Search31.grid1.TextMatrix(Search31.grid1.Row, 0)
    Unload Search31
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
Private Sub LoadPhoto(sModelNo As String)
On Error GoTo myerror
ImgX1.Images.Clear
ImgX1.Import.FromFile retFile(sModelNo)
Exit Sub
myerror:
Err.Clear
End Sub

Private Sub ImgX1_Click()
If validPhoto(retFile(grid1.TextMatrix(grid1.Row, 0))) Then
    Dim oMaxPhoto As New maxPhoto
    oMaxPhoto.cFile = retFile(grid1.TextMatrix(grid1.Row, 0))
    oMaxPhoto.cDesca = grid1.TextMatrix(grid1.Row, 4)
    oMaxPhoto.Show 1
End If
End Sub
Sub DocDisc_Look()
    Dim Generalarray(5)
    Dim listarray(0, 4)
    Dim GrdArray(6, 1)
    
    Set Generalarray(0) = Me
    Generalarray(1) = "SELECT DOC_NO,DATE, CONVERT(VARCHAR(10),[DATE],111),FILE0_90H.DESCA , isstop , date1 , date2 " & _
                      " FROM FILE0_90H "
    If cFilter <> "" Then
        Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
    End If
    
    Generalarray(2) = "Order by Date , DOC_NO "
    Generalarray(3) = 4200
    Generalarray(5) = False
    
    
    listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ-«·»Ì«‰"
    listarray(0, 1) = "(@@Doc_No@@6 OR " & _
                      " ##[DATE]## or %%FILE0_90h.DESCA%%)"
    
    GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
    GrdArray(0, 1) = 1000
    
    GrdArray(1, 0) = "«· «—ÌŒ"
    GrdArray(1, 1) = 0
    
    GrdArray(2, 0) = "«· «—ÌŒ"
    GrdArray(2, 1) = 1500
    
    GrdArray(3, 0) = "«·»Ì«‰"
    GrdArray(3, 1) = 4000
    
    GrdArray(4, 0) = "„ Êﬁ›"
    GrdArray(4, 1) = 800
    
    GrdArray(5, 0) = "„‰  «—ÌŒ"
    GrdArray(5, 1) = 1500
    
    GrdArray(6, 0) = "≈·Ï"
    GrdArray(6, 1) = 1500
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Load Search31
    Search31.Caption = "«” ⁄·«„"
    Search31.Show 1
End Sub
Private Sub myload9()
Dim i As Double
Dim ScalTable As New ADODB.Recordset
Dim cStrScal As String

Dim cString  As String, cWhere As String
'                           0               1                 2                3               4                5
    
    cStrScal = " SELECT  FILE1_10.SCAL FROM FILE1_10 INNER JOIN QFILE6_20 ON FILE1_10.ITEM = QFILE6_20.ITEM" & _
              " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
              " inner join file4_10 on file4_10.code = file1_10.code" & _
              " INNER JOIN STORE_BR ON QFILE6_20.STORE = STORE_BR.CODE"

    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    If xcust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.CUST = " & MyParn(xcust.BoundText)
    If xMosm.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    
    
    If xmosm22.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
    If xsex.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xage.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    
    If XDOC_DISC.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.DOC_DISC = " & MyParn(XDOC_DISC.BoundText)
    
    If XOutlet.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " file1_10.Outlet = 1 "
    
    If xSupp.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGrCust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
    If xGroup.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    
    
    If xdoc_no.text <> "" Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.doc_no = " & MyParn(xdoc_no.text)
    If XSTORE.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " STORE = " & MyParn(XSTORE.BoundText)
    If xSection.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[Section] = " & MyParn(xSection.BoundText)
    If IsDate(xDate1.text) Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.[DATE] <= " & DateSq(xdate2.text)
    
    If xall(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.ISBRANCH = 0 "
    
    If xonest(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 0 "
    If xonest(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 1 "
    If XISONLINE.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.online <> 0 "
    If XSHIP.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.[SHIP] = " & MyParn(XSHIP.BoundText)
    
    cStrScal = cStrScal & " GROUP BY file1_10.SCAL" & _
                          " ORDER BY file1_10.SCAL"
    
    Set ScalTable = mycmd(cStrScal, con)
    
    cString = "SELECT ' ' AS N1 , FACT.desca , file1_10.ModelNo , file1_10.mosm , file1_10.MODELFACT0, file1_10.desca ,   " & _
        " SUM(QFILE6_20.QUANT) AS t_q "
    
    With ScalTable
        Do While Not ScalTable.EOF
            cString = cString & "," & myiif("FILE1_10.SCAL = " & MyParn(ScalTable!SCAL), "QFILE6_20.QUANT") & " AS '" & ScalTable!SCAL & "'"
            .MoveNext
        Loop
    End With
    cString = cString & " FROM FILE1_10 INNER JOIN QFILE6_20 ON FILE1_10.ITEM = QFILE6_20.ITEM" & _
                        " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
                        " inner join file4_10 on file4_10.code = file1_10.code" & _
                        " INNER JOIN STORE_BR ON QFILE6_20.STORE = STORE_BR.CODE"
                        
    If cWhere <> "" Then cString = cString & " WHERE " & cWhere
    
    cString = cString & " GROUP BY file1_10.ModelNo, FACT.desca , file1_10.mosm, file1_10.MODELFACT0, file1_10.desca" & _
                         "  ORDER BY file1_10.ModelNo "
    Set data29.Recordset = mycmd(cString, con)
    XTEXT1.text = cString
    
    FIXGRID9
'If grid1.Rows > 0 Then ' grid1.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FIXGRID9()
With grid9
.RowHeight(0) = 900
.WordWrap = True
.TextMatrix(0, 0) = "«·„’‰⁄"
.TextMatrix(0, 1) = "„ÊœÌ·"
.TextMatrix(0, 2) = "„ÊœÌ·"
.TextMatrix(0, 3) = "«·’‰›"

.TextMatrix(0, 4) = "„Ê”„"
.TextMatrix(0, 5) = "„ﬁ«”"
.TextMatrix(0, 6) = " »«—ﬂÊœ Œ«—ÃÏ"
.TextMatrix(0, 7) = " »«—ﬂÊœ „’‰⁄"
.TextMatrix(0, 8) = " »«—ﬂÊœ 13"
.TextMatrix(0, 9) = "BARCODE GS1"
.TextMatrix(0, 10) = "⁄œœ „»Ì⁄« "

.ColWidth(0) = 2500
.ColWidth(1) = 2500
.ColWidth(2) = 1000
.ColWidth(3) = 3000
.ColWidth(4) = 1200
.ColWidth(5) = 1000
.ColWidth(6) = 2000
.ColWidth(7) = 1500
.ColWidth(8) = 2000
.ColWidth(9) = 2100
.ColWidth(10) = 1400

.MergeCol(0) = True
.MergeCol(1) = True
.MergeCol(2) = True
.MergeCol(3) = True
.MergeCol(4) = True

.ColHidden(1) = True

.MergeCells = flexMergeRestrictRows


For col = 10 To .Cols - 1
    .ColDataType(col) = flexDTDouble
    .Subtotal flexSTSum, -1, col, "#0", vbYellow, , True, ""
Next
.ExplorerBar = flexExSort
.Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
End With
End Sub
Private Sub myloadGrd9()

Dim cString  As String
Dim cWhere As String
    
If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
If xcust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.CUST = " & MyParn(xcust.BoundText)
If xMosm.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)

If XMOSM2.text <> "W" Then
    If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
Else
    If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
End If


If xmosm22.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
If xsex.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
If xage.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)

If XDOC_DISC.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.DOC_DISC = " & MyParn(XDOC_DISC.BoundText)

If XOutlet.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " file1_10.Outlet = 1 "

If xSupp.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
If xGrCust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
If xGroup.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
If xFact.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)

If Not bOpt5 Then
    cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
End If
     
If xdoc_no.text <> "" Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.doc_no = " & MyParn(xdoc_no.text)
If XSTORE.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " STORE = " & MyParn(XSTORE.BoundText)
If xSection.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[Section] = " & MyParn(xSection.BoundText)
If IsDate(xDate1.text) Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.[DATE] >= " & DateSq(xDate1.text)
If IsDate(xdate2.text) Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.[DATE] <= " & DateSq(xdate2.text)

If xall(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.ISBRANCH = 1 "
If xall(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.ISBRANCH = 0 "

If xonest(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 0 "
If xonest(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 1 "
If XISONLINE.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.online <> 0 "
If XSHIP.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.[SHIP] = " & MyParn(XSHIP.BoundText)
    

    cString = "SELECT FACT.desca," & _
              "file1_10.ModelNo," & _
              "file1_10.MODELFACT0," & _
              "file1_10.desca," & _
              "file1_10.mosm," & _
              "FILE1_10.SCAL," & _
              "FILE1_10.BARCODE," & _
              "FILE1_10.BARCODE2," & _
              "FILE1_10.BARCODE13," & _
              "FILE1_10.BARCODE_GS1," & _
              "SUM(VW_FILE6_20.QUANT) AS t_q " & _
              " FROM FILE1_10 " & _
              " INNER JOIN VW_FILE6_20 ON FILE1_10.ITEM = VW_FILE6_20.ITEM" & _
              " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
              " inner join file4_10 on file4_10.code = file1_10.code" & _
              " INNER JOIN STORE_BR ON VW_FILE6_20.STORE = STORE_BR.CODE"
                        
    If cWhere <> "" Then cString = cString & " WHERE " & cWhere
    
    cString = cString & _
              " GROUP BY FACT.desca," & _
              "file1_10.ModelNo," & _
              "file1_10.mosm," & _
              "file1_10.MODELFACT0," & _
              "file1_10.desca," & _
              "FILE1_10.SCAL," & _
              "FILE1_10.BARCODE," & _
              "FILE1_10.BARCODE2," & _
              "FILE1_10.BARCODE13," & _
              "FILE1_10.BARCODE_GS1" & _
              " ORDER BY file1_10.ModelNo," & _
              " FILE1_10.SCAL"
    Set data29.Recordset = cmd(cString, con).Execute
    
    XTEXT1.text = cString
    
    FIXGRID9
'If grid1.Rows > 0 Then ' grid1.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FIXGRD9()
With grid9
    .RowHeight(0) = 900
    .WordWrap = True
    .TextMatrix(0, 0) = "«·„’‰⁄"
    .TextMatrix(0, 1) = "„ÊœÌ·"
    .TextMatrix(0, 2) = "„ÊœÌ·"
    .TextMatrix(0, 3) = "«·’‰›"
    
    .TextMatrix(0, 4) = "„Ê”„"
    .TextMatrix(0, 5) = "„ﬁ«”"
    .TextMatrix(0, 6) = " »«—ﬂÊœ Œ«—ÃÏ"
    .TextMatrix(0, 7) = " »«—ﬂÊœ „’‰⁄"
    .TextMatrix(0, 8) = " »«—ﬂÊœ 13"
    .TextMatrix(0, 9) = "BARCODE GS1"
    .TextMatrix(0, 10) = "⁄œœ „»Ì⁄« "
    
    .ColWidth(0) = 2500
    .ColWidth(1) = 2500
    .ColWidth(2) = 1000
    .ColWidth(3) = 3000
    .ColWidth(4) = 1200
    .ColWidth(5) = 1000
    .ColWidth(6) = 2000
    .ColWidth(7) = 1500
    .ColWidth(8) = 2000
    .ColWidth(9) = 2100
    .ColWidth(10) = 1400
    
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    
    .ColHidden(1) = True
    
    .MergeCells = flexMergeRestrictRows
    
    
    For col = 10 To .Cols - 1
        .ColDataType(col) = flexDTDouble
        .Subtotal flexSTSum, -1, col, "#0", vbYellow, , True, ""
    Next
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    End With
End Sub
Private Sub myload8()
Dim i As Double
Dim cStrScal As String
Dim ScalTable As New ADODB.Recordset
Dim cString  As String, cWhere As String, cString2 As String
'                           0               1                 2                3               4                5
    
    cStrScal = " SELECT  FILE1_10.SCAL " & _
                " FROM FILE1_10  INNER JOIN QFILE6_20 ON FILE1_10.ITEM = QFILE6_20.ITEM " & _
                " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
                " INNER JOIN FILE4_10 ON FILE4_10.CODE = FILE1_10.CODE" & _
                " INNER JOIN STORE_BR ON QFILE6_20.STORE = STORE_BR.CODE"
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    
    If xcust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.CUST = " & MyParn(xcust.BoundText)
    If xMosm.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    
    If xmosm22.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
    If xsex.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xage.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If XDOC_DISC.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.DOC_DISC = " & MyParn(XDOC_DISC.BoundText)
    If XOutlet.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " file1_10.Outlet = 1 "
    If xSupp.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGrCust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
    If xGroup.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If xdoc_no.text <> "" Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.doc_no = " & MyParn(xdoc_no.text)
    If XSTORE.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " STORE = " & MyParn(XSTORE.BoundText)
    If xSection.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[Section] = " & MyParn(xSection.BoundText)
    If IsDate(xDate1.text) Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.[DATE] <= " & DateSq(xdate2.text)
    If xall(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.ISBRANCH = 0 "
    If xonest(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 0 "
    If xonest(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 1 "
    If XISONLINE.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.online <> 0 "
    If XSHIP.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " QFILE6_20.[SHIP] = " & MyParn(XSHIP.BoundText)
    
  
    If cWhere <> "" Then cStrScal = cStrScal & " WHERE " & cWhere
    cStrScal = cStrScal & " GROUP BY file1_10.SCAL" & _
                          " ORDER BY file1_10.SCAL"
    
    Set ScalTable = mycmd(cStrScal, con)
    
    
    cString = " SELECT FILE1_50.CODE , FILE1_50.DESCA ,   " & _
              " SUM(QFILE6_20.QUANT) AS t_q "
    With ScalTable
        Do While Not ScalTable.EOF
            cString = cString & "," & myiif("FILE1_10.SCAL = " & MyParn(ScalTable!SCAL), "QFILE6_20.QUANT") & " AS '" & ScalTable!SCAL & "'"
            .MoveNext
        Loop
    End With

    cString = cString & " FROM FILE1_10 INNER JOIN QFILE6_20 ON FILE1_10.ITEM = QFILE6_20.ITEM " & _
                        " INNER JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE" & _
                        " inner join file4_10 on file4_10.code = file1_10.code" & _
                        " INNER JOIN STORE_BR ON QFILE6_20.STORE = STORE_BR.CODE"
                        
    
    If cWhere <> "" Then cString = cString & " WHERE " & cWhere
    cString = cString & " GROUP BY file1_50.CODE , FILE1_50.desca" & _
                        " ORDER BY file1_50.CODE , FILE1_50.desca  "
    
    Set data28.Recordset = mycmd(cString, con)
    XTEXT1.text = cString
    
    FIXGRID8
End Sub
Sub FIXGRID8()
With grid8
    .RowHeight(0) = 600
    .WordWrap = True
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·„Ã„Ê⁄…"
    .TextMatrix(0, 2) = "⁄œœ „»Ì⁄« "

    
    .ColWidth(0) = 600
    .ColWidth(1) = 4000
    .ColWidth(2) = 1000
    For nCol = 2 To .Cols - 1
        .ColWidth(nCol) = 700
        .ColDataType(nCol) = flexDTDouble
        .Subtotal flexSTSum, -1, nCol, "#0", vbYellow, , True, ""
    Next nCol
    .Subtotal flexSTSum, -1, 2, "#0", vbYellow, , True, ""
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    End With
End Sub
Private Sub myload7()
Dim i As Double
Dim cString  As String, cWhere As String '                           0               1                 2                3               4                5
If SEC(0).Value Then
    cString = " SELECT FILE4_10.[GROUP]," & _
                " FILE4_10.desca," & _
                "file1_10.mosm," & _
                " SUM(vw_FILE6_20.QUANT) AS t_q, " & _
                " SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET, " & _
                " SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
                " SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft ," & _
                " ROUND(CASE WHEN SUM(vw_FILE6_20.TOTAL_NET) <> 0 THEN SUM(vw_FILE6_20.ITEM_PROFIT)/SUM(vw_FILE6_20.TOTAL_NET) ELSE 0 END * 100,2)," & _
                " ROUND(CASE WHEN SUM(SUM(vw_FILE6_20.TOTAL_NET)) OVER () <> 0 THEN SUM(vw_FILE6_20.TOTAL_NET)/SUM(SUM(vw_FILE6_20.TOTAL_NET)) OVER () ELSE 0 END * 100,2)" & _
                " FROM FILE1_10 " & _
                " INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM" & _
                " INNER JOIN FILE4_10 ON FILE1_10.CODE = FILE4_10.CODE" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"
ElseIf SEC(1).Value Then
    cString = "SELECT FILE4_10.[GROUP]," & _
                "FILE4_10.desca," & _
                "file0_81.DESCA," & _
                " SUM(vw_FILE6_20.QUANT) AS t_q, " & _
                " SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET, " & _
                " SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
                " SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft ," & _
                " ROUND(CASE WHEN SUM(vw_FILE6_20.TOTAL_NET) <> 0 THEN SUM(vw_FILE6_20.ITEM_PROFIT)/SUM(vw_FILE6_20.TOTAL_NET) ELSE 0 END * 100,2)," & _
                " ROUND(CASE WHEN SUM(SUM(vw_FILE6_20.TOTAL_NET)) OVER () <> 0 THEN SUM(vw_FILE6_20.TOTAL_NET)/SUM(SUM(vw_FILE6_20.TOTAL_NET)) OVER () ELSE 0 END * 100,2)" & _
                " FROM FILE1_10 INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM" & _
                " INNER JOIN FILE4_10 ON FILE1_10.CODE = FILE4_10.CODE" & _
                " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"

ElseIf SEC(2).Value Then
    cString = "SELECT FILE4_10.[GROUP]," & _
                "FILE4_10.desca," & _
                "file0_82.DESCA," & _
                " SUM(vw_FILE6_20.QUANT) AS t_q, " & _
                " SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET, " & _
                " SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
                " SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft ," & _
                " ROUND(CASE WHEN SUM(vw_FILE6_20.TOTAL_NET) <> 0 THEN SUM(vw_FILE6_20.ITEM_PROFIT)/SUM(vw_FILE6_20.TOTAL_NET) ELSE 0 END * 100,2)," & _
                " ROUND(CASE WHEN SUM(SUM(vw_FILE6_20.TOTAL_NET)) OVER () <> 0 THEN SUM(vw_FILE6_20.TOTAL_NET)/SUM(SUM(vw_FILE6_20.TOTAL_NET)) OVER () ELSE 0 END * 100,2)" & _
                " FROM FILE1_10 INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM" & _
                " INNER JOIN FILE4_10 ON FILE1_10.CODE = FILE4_10.CODE" & _
                " LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"

ElseIf SEC(3).Value Then
    cString = "SELECT FILE4_10.[GROUP]," & _
                "FILE4_10.desca," & _
                "NULL," & _
                " SUM(vw_FILE6_20.QUANT) AS t_q, " & _
                " SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET, " & _
                " SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
                " SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft ," & _
                " ROUND(CASE WHEN SUM(vw_FILE6_20.TOTAL_NET) <> 0 THEN SUM(vw_FILE6_20.ITEM_PROFIT)/SUM(vw_FILE6_20.TOTAL_NET) ELSE 0 END * 100,2)," & _
                " ROUND(CASE WHEN SUM(SUM(vw_FILE6_20.TOTAL_NET)) OVER () <> 0 THEN SUM(vw_FILE6_20.TOTAL_NET)/SUM(SUM(vw_FILE6_20.TOTAL_NET)) OVER () ELSE 0 END * 100,2)" & _
                " FROM FILE1_10 INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM" & _
                " INNER JOIN FILE4_10 ON FILE1_10.CODE = FILE4_10.CODE" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"
End If

If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    
If xcust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.CUST = " & MyParn(xcust.BoundText)
If xMosm.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)

If XMOSM2.text <> "W" Then
    If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
Else
    If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
End If

If xmosm22.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
If xsex.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
If xage.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
If XDOC_DISC.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.DOC_DISC = " & MyParn(XDOC_DISC.BoundText)
If XOutlet.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " file1_10.Outlet = 1 "
If xSupp.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
If xGrCust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
If xGroup.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
If xFact.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
If xdoc_no.text <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.doc_no = " & MyParn(xdoc_no.text)
If XSTORE.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " STORE = " & MyParn(XSTORE.BoundText)
If xSection.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[Section] = " & MyParn(xSection.BoundText)
If IsDate(xDate1.text) Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[DATE] >= " & DateSq(xDate1.text)
If IsDate(xdate2.text) Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[DATE] <= " & DateSq(xdate2.text)
If xall(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.ISBRANCH = 1 "
If xall(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.ISBRANCH = 0 "
If xonest(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 0 "
If xonest(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 1 "
If XISONLINE.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.online <> 0 "
If XSHIP.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[SHIP] = " & MyParn(XSHIP.BoundText)

If cWhere <> "" Then cString = cString & " WHERE " & cWhere

If SEC(0).Value <> 0 Then
    cString = cString & " GROUP BY FILE4_10.[GROUP] , FILE4_10.CODE,FILE4_10.DESCA , file1_10.mosm ORDER BY FILE4_10.[GROUP], FILE4_10.CODE  "
End If
If SEC(1).Value <> 0 Then
    cString = cString & " GROUP BY FILE4_10.[GROUP] , FILE4_10.CODE, FILE4_10.desca , file0_81.DESCA ORDER BY FILE4_10.[GROUP], FILE4_10.CODE  "
End If
If SEC(2).Value <> 0 Then
    cString = cString & " GROUP BY FILE4_10.[GROUP] , FILE4_10.CODE, FILE4_10.desca , file0_82.DESCA ORDER BY FILE4_10.[GROUP], FILE4_10.CODE  "
End If
If SEC(3).Value <> 0 Then
    cString = cString & " GROUP BY FILE4_10.[GROUP] , FILE4_10.CODE, FILE4_10.desca ORDER BY FILE4_10.[GROUP], FILE4_10.CODE  "
End If
Set data27.Recordset = mycmd(cString, con)
XTEXT1.text = cString

FIXGRID7
End Sub
Sub FIXGRID7()
With grid7
    .FixedRows = 1
    .RowHeight(0) = 600
    .WordWrap = True
    .Cols = 9
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·„Ê—œ"
    If SEC(0).Value <> 0 Then .TextMatrix(0, 2) = "«·„Ê”„"
    If SEC(1).Value <> 0 Then .TextMatrix(0, 2) = "«·‰Ê⁄"
    If SEC(2).Value <> 0 Then .TextMatrix(0, 2) = "«·›∆… «·⁄„—Ì… "
    If SEC(3).Value <> 0 Then .ColHidden(2) = True
    
    .TextMatrix(0, 3) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 4) = "ﬁÌ„… „»Ì⁄« "
    .TextMatrix(0, 5) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 6) = "—»Õ „»Ì⁄« "
    .TextMatrix(0, 7) = "‰”»… —»Õ"
    .TextMatrix(0, 8) = "‰”»… «·„Ê—œ"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 2500
    .ColWidth(2) = 1000
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    .ColWidth(7) = 1500
    .ColWidth(8) = 1500
    
    .ColHidden(5) = Not bOpt10
    .ColHidden(6) = Not bOpt10
    .ColHidden(7) = Not bOpt10
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    For nRow = 1 To .Rows - 1
'        .TextMatrix(nRow, 6) = Round(Val(.TextMatrix(nRow, 4)) - Val(.TextMatrix(nRow, 5)), 2)
'        npro = 0
'        If Val(.TextMatrix(nRow, 5)) <> 0 Then npro = (Val(.TextMatrix(nRow, 6)) / Val(.TextMatrix(nRow, 5)) * 100)
'        .TextMatrix(nRow, 7) = Round(npro, 2)
    Next nRow
    For nCol = 3 To .Cols - 1
'        .ColFormat(nCol) = "#0.00"
'        .ColDataType(nCol) = flexDTDouble
'        For nRow = 1 To .Rows - 1
'            .TextMatrix(nRow, nCol) = Format(.TextMatrix(nRow, nCol), "#0.00")
'        Next nRow
    Next nCol
    
    .SubtotalPosition = flexSTAbove
    For i = 3 To .Cols - 3
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
        '.Subtotal flexSTSum, 0, i, "#0", vbRed, vbYellow, True, "  "
    Next i
    If .Rows > 2 Then
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
        .TextMatrix(1, 7) = Round((.ValueMatrix(1, 6) / .ValueMatrix(1, 4)) * 100, 2)
    End If
    For nRow = 1 To .Rows - 1
'        nRate = 0
'        nTot = Val(.TextMatrix(.Rows - 1, 4))
'        If nTot <> 0 Then .TextMatrix(nRow, 8) = Round(Val(.TextMatrix(nRow, 4)) / nTot * 100, 2)
    Next nRow
'    .Subtotal flexSTSum, -1, 8, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, 0, 8, "#0", vbRed, vbYellow, True, "  "
    End With
End Sub
Private Sub myload6()
Dim i As Double
Dim cString  As String, cWhere As String
'                           0               1                 2                3               4                5
    
    If SEC(0).Value Then
        cString = " SELECT FACT.[CODE]," & _
                "  FACT.desca," & _
                "  file1_10.mosm ,   " & _
                "  SUM(vw_FILE6_20.QUANT) AS t_q, " & _
                "  SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET, " & _
                "  SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
                "  SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft ," & _
                " ROUND(CASE WHEN SUM(vw_FILE6_20.TOTAL_NET) <> 0 THEN SUM(vw_FILE6_20.ITEM_PROFIT)/SUM(vw_FILE6_20.TOTAL_NET) ELSE 0 END * 100,2)," & _
                " ROUND(CASE WHEN SUM(SUM(vw_FILE6_20.TOTAL_NET)) OVER () <> 0 THEN SUM(vw_FILE6_20.TOTAL_NET)/SUM(SUM(vw_FILE6_20.TOTAL_NET)) OVER () ELSE 0 END * 100,2)" & _
                " FROM FILE1_10 INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM" & _
                " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
                " inner join file4_10 on file4_10.code = file1_10.code" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"
    ElseIf SEC(1).Value Then
        cString = " SELECT FACT.[CODE]," & _
                  "  FACT.desca," & _
                  "file0_81.DESCA ,   " & _
                "  SUM(vw_FILE6_20.QUANT) AS t_q, " & _
                "  SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET, " & _
                "  SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
                "  SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft ," & _
                " ROUND(CASE WHEN SUM(vw_FILE6_20.TOTAL_NET) <> 0 THEN SUM(vw_FILE6_20.ITEM_PROFIT)/SUM(vw_FILE6_20.TOTAL_NET) ELSE 0 END * 100,2)," & _
                " ROUND(CASE WHEN SUM(SUM(vw_FILE6_20.TOTAL_NET)) OVER () <> 0 THEN SUM(vw_FILE6_20.TOTAL_NET)/SUM(SUM(vw_FILE6_20.TOTAL_NET)) OVER () ELSE 0 END * 100,2)" & _
                " FROM FILE1_10 INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM" & _
                " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
                " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX" & _
                " inner join file4_10 on file4_10.code = file1_10.code" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"
    ElseIf SEC(2).Value Then
        cString = " SELECT FACT.[CODE]," & _
                "FACT.desca," & _
                "file0_82.DESCA," & _
                "  SUM(vw_FILE6_20.QUANT) AS t_q, " & _
                "  SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET, " & _
                "  SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
                "  SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft ," & _
                " ROUND(CASE WHEN SUM(vw_FILE6_20.TOTAL_NET) <> 0 THEN SUM(vw_FILE6_20.ITEM_PROFIT)/SUM(vw_FILE6_20.TOTAL_NET) ELSE 0 END * 100,2)," & _
                " ROUND(CASE WHEN SUM(SUM(vw_FILE6_20.TOTAL_NET)) OVER () <> 0 THEN SUM(vw_FILE6_20.TOTAL_NET)/SUM(SUM(vw_FILE6_20.TOTAL_NET)) OVER () ELSE 0 END * 100,2)" & _
                " FROM FILE1_10 INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM" & _
                " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
                " LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE" & _
                " inner join file4_10 on file4_10.code = file1_10.code" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"
    ElseIf SEC(3).Value <> 0 Then
        cString = " SELECT FACT.[CODE]," & _
                "FACT.desca ," & _
                "NULL ," & _
                "  SUM(vw_FILE6_20.QUANT) AS t_q, " & _
                "  SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET, " & _
                "  SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
                "  SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft ," & _
                " ROUND(CASE WHEN SUM(vw_FILE6_20.TOTAL_NET) <> 0 THEN SUM(vw_FILE6_20.ITEM_PROFIT)/SUM(vw_FILE6_20.TOTAL_NET) ELSE 0 END * 100,2)," & _
                " ROUND(CASE WHEN SUM(SUM(vw_FILE6_20.TOTAL_NET)) OVER () <> 0 THEN SUM(vw_FILE6_20.TOTAL_NET)/SUM(SUM(vw_FILE6_20.TOTAL_NET)) OVER () ELSE 0 END * 100,2)" & _
                " FROM FILE1_10 INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM" & _
                " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"
    End If

    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    
    If xcust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.CUST = " & MyParn(xcust.BoundText)
    If xMosm.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    
    If xmosm22.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
    If xsex.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xage.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If XDOC_DISC.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.DOC_DISC = " & MyParn(XDOC_DISC.BoundText)
    If XOutlet.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " file1_10.Outlet = 1 "
    If xSupp.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGrCust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
    If xGroup.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If xdoc_no.text <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.doc_no = " & MyParn(xdoc_no.text)
    If XSTORE.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " STORE = " & MyParn(XSTORE.BoundText)
    If xSection.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[Section] = " & MyParn(xSection.BoundText)
    If IsDate(xDate1.text) Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[DATE] <= " & DateSq(xdate2.text)
    If xall(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.ISBRANCH = 0 "
    If xonest(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 0 "
    If xonest(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 1 "
    If XSHIP.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[SHIP] = " & MyParn(XSHIP.BoundText)
    If XISONLINE.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.online <> 0 "
    If XSHIP.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[SHIP] = " & MyParn(XSHIP.BoundText)
    


    If cWhere <> "" Then cString = cString & " WHERE " & cWhere

    If SEC(0).Value <> 0 Then
        cString = cString & " GROUP BY FACT.[CODE] ,FACT.DESCA , file1_10.mosm ORDER BY FACT.[CODE] ,FACT.DESCA  "
    End If
    If SEC(1).Value <> 0 Then
        cString = cString & " GROUP BY FACT.[CODE] ,FACT.DESCA  , file0_81.DESCA ORDER BY FACT.[CODE] ,FACT.DESCA  "
    End If
    If SEC(2).Value <> 0 Then
        cString = cString & " GROUP BY FACT.[CODE] ,FACT.DESCA  , file0_82.DESCA ORDER BY FACT.[CODE] ,FACT.DESCA  "
    End If
    If SEC(3).Value <> 0 Then
        cString = cString & " GROUP BY FACT.[CODE] ,FACT.DESCA  ORDER BY FACT.[CODE] ,FACT.DESCA  "
    End If
    
    Set data26.Recordset = myRecordSet(cString, con)
    XTEXT1.text = cString
    
    FIXGRID6
End Sub
Sub FIXGRID6()
With grid6
    .RowHeight(0) = 600
    .Cols = 9
    .FixedRows = 1
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    If SEC(0).Value <> 0 Then .TextMatrix(0, 2) = "«·„Ê”„"
    If SEC(1).Value <> 0 Then .TextMatrix(0, 2) = "«·‰Ê⁄"
    If SEC(2).Value <> 0 Then .TextMatrix(0, 2) = "«·›∆… «·⁄„—Ì… "
    If SEC(3).Value <> 0 Then .ColHidden(2) = True
    
    .TextMatrix(0, 3) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 4) = "ﬁÌ„… „»Ì⁄« "
    
    .TextMatrix(0, 5) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 6) = "—»Õ „»Ì⁄« "
    .TextMatrix(0, 7) = "‰”»… —»Õ"
    .TextMatrix(0, 8) = "‰”»… «·„’‰⁄"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 2500
    .ColWidth(2) = 1000
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    .ColWidth(7) = 1500
    .ColWidth(8) = 1500
    
    .ColHidden(5) = Not bOpt10
    .ColHidden(6) = Not bOpt10
    .ColHidden(7) = Not bOpt10
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4

     For nRow = 1 To .Rows - 1
            '.TextMatrix(nRow, 6) = Round(Val(.TextMatrix(nRow, 4)) - Val(.TextMatrix(nRow, 5)), 2)
            'npro = 0
            'If Val(.TextMatrix(nRow, 5)) <> 0 Then npro = (Val(.TextMatrix(nRow, 6)) / Val(.TextMatrix(nRow, 5)) * 100)
'            .TextMatrix(nRow, 7) = Round(npro, 2)
     Next nRow
    
    For nCol = 3 To .Cols - 1
        .ColFormat(nCol) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
'        For nRow = 1 To .Rows - 1
'            .TextMatrix(nRow, nCol) = Format(.TextMatrix(nRow, nCol), "#0.00")
'        Next nRow
    Next nCol
    .SubtotalPosition = flexSTAbove
    For i = 3 To .Cols - 3
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
'       .Subtotal flexSTSum, 0, i, "#0", vbRed, vbYellow, True, "  "
    Next i
    
    If .Rows > 2 Then
        .Cell(flexcpAlignment, 1, 0, 1, 1) = 7
        If .ValueMatrix(1, 4) <> 0 Then
           .TextMatrix(1, 7) = (.ValueMatrix(1, 6) / .ValueMatrix(1, 4)) * 100
        End If
    End If
    
    For nRow = 1 To .Rows - 1
        'nRate = 0
        'nTot = Val(.TextMatrix(.Rows - 1, 4))
        'If nTot <> 0 Then .TextMatrix(nRow, 8) = Round(Val(.TextMatrix(nRow, 4)) / nTot * 100, 2)
    Next nRow
'    .Subtotal flexSTSum, -1, 8, "#0", vbRed, vbYellow, True, "  "
'   .Subtotal flexSTSum, 0, 8, "#0", vbRed, vbYellow, True, "  "
    End With
End Sub
Private Sub myload5()
Dim i As Double
Dim cString  As String, cWhere As String
'                           0               1                 2                3               4                5
    
    
    If SEC(0).Value <> 0 Then
        cString = " SELECT FILE1_10.[SECTION]," & _
                  " NULL as n2," & _
                  "FILE1_10.MOSM,   " & _
                "  SUM(vw_FILE6_20.QUANT) AS t_q, " & _
                "  SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET, " & _
                "  SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
                "  SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft " & _
                " FROM FILE1_10 " & _
                " INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM" & _
                " inner join file4_10 on file4_10.code = file1_10.code" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"

    ElseIf SEC(1).Value <> 0 Then
        cString = " SELECT FILE1_10.[SECTION]," & _
                  " NULL as n2," & _
                  "FILE0_81.desca ,   " & _
                "  SUM(vw_FILE6_20.QUANT) AS t_q, " & _
                "  SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET, " & _
                "  SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
                "  SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft " & _
                " FROM FILE1_10 " & _
                " INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM " & _
                " inner join file0_81 on file0_81.code = file1_10.modelsex" & _
                " inner join file4_10 on file4_10.code = file1_10.code" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"
    ElseIf SEC(2).Value <> 0 Then
        cString = " SELECT FILE1_10.[SECTION]," & _
                  " NULL as n2," & _
                  "FILE0_82.desca ,   " & _
                "  SUM(vw_FILE6_20.QUANT) AS t_q, " & _
                "  SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET, " & _
                "  SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
                "  SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft " & _
                " FROM FILE1_10 " & _
                " INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM" & _
                " inner join file0_82 on file0_82.code = file1_10.modelage" & _
                " inner join file4_10 on file4_10.code = file1_10.code" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"
    ElseIf SEC(3).Value <> 0 Then
        cString = " SELECT FILE1_10.[SECTION]," & _
                  " NULL as n2," & _
                  "NULL,   " & _
                "  SUM(vw_FILE6_20.QUANT) AS t_q, " & _
                "  SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET, " & _
                "  SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
                "  SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft " & _
                " FROM FILE1_10 " & _
                " INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM" & _
                " inner join file4_10 on file4_10.code = file1_10.code" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"
    End If
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"

    If xcust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.CUST = " & MyParn(xcust.BoundText)
    If xMosm.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    
    If xmosm22.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
    If xsex.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xage.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If XDOC_DISC.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.DOC_DISC = " & MyParn(XDOC_DISC.BoundText)
    If XOutlet.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " file1_10.Outlet = 1 "
    If xSupp.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGrCust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
    If xGroup.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If xdoc_no.text <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.doc_no = " & MyParn(xdoc_no.text)
    If XSTORE.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " STORE = " & MyParn(XSTORE.BoundText)
    If xSection.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[Section] = " & MyParn(xSection.BoundText)
    If IsDate(xDate1.text) Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[DATE] <= " & DateSq(xdate2.text)
    If xall(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.ISBRANCH = 0 "
    If xonest(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 0 "
    If xonest(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 1 "
    If XISONLINE.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.online <> 0 "
    If XSHIP.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[SHIP] = " & MyParn(XSHIP.BoundText)


    If cWhere <> "" Then cString = cString & " WHERE " & cWhere

    If SEC(0).Value <> 0 Then cString = cString & " GROUP BY FILE1_10.[SECTION] , FILE1_10.MOSM     ORDER BY FILE1_10.[SECTION] "
    If SEC(1).Value <> 0 Then cString = cString & " GROUP BY FILE1_10.[SECTION] , FILE0_81.desca    ORDER BY FILE1_10.[SECTION] "
    If SEC(2).Value <> 0 Then cString = cString & " GROUP BY FILE1_10.[SECTION] , FILE0_82.desca    ORDER BY FILE1_10.[SECTION] "
    If SEC(3).Value <> 0 Then cString = cString & " GROUP BY FILE1_10.[SECTION] ORDER BY FILE1_10.[SECTION] "
    Set data25.Recordset = mycmd(cString, con)
    XTEXT1.text = cString
    
    FIXGRID5
End Sub
Sub FIXGRID5()
With grid5
    .RowHeight(0) = 600
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·ﬁ”„"
    
    If SEC(0).Value <> 0 Then .TextMatrix(0, 2) = "«·„Ê”„"
    If SEC(1).Value <> 0 Then .TextMatrix(0, 2) = "«·‰Ê⁄"
    If SEC(2).Value <> 0 Then .TextMatrix(0, 2) = "«·›∆… «·⁄„—Ì…"
    If SEC(3).Value <> 0 Then .ColHidden(2) = True
    
    .TextMatrix(0, 3) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 4) = "ﬁÌ„… „»Ì⁄« "
    
    .TextMatrix(0, 5) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 6) = "—»Õ „»Ì⁄« "
    
    .ColWidth(0) = 500
    .ColWidth(1) = 2500
    .ColWidth(2) = 1000
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    
    .ColHidden(5) = Not bOpt10
    .ColHidden(6) = Not bOpt10
    
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
   
    For nRow = 1 To .Rows - 1
        '.TextMatrix(nRow, 6) = Round(Val(.TextMatrix(nRow, 4)) - Val(.TextMatrix(nRow, 5)), 2)
        If .TextMatrix(nRow, 0) <> "" Then .TextMatrix(nRow, 1) = GetDesca("SELECT DESCA FROM FILE1_10SC WHERE CODE = " & MyParn(.TextMatrix(nRow, 0)), con)
    Next nRow
    For nCol = 3 To .Cols - 1
        .ColFormat(nCol) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol) = Format(.TextMatrix(nRow, nCol), "#0.00")
        Next nRow
    Next nCol
    
    .SubtotalPosition = flexSTAbove
    For i = 3 To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next i
    If .Rows > 1 Then .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub

Private Sub MYLOAD4()
Dim i As Double
Dim cString  As String, cWhere As String
    If SEC(0).Value Then
        cString = " SELECT FILE1_10.[GROUP]," & _
                "   FILE1_50.DESCA," & _
                " File1_10.mosm, " & _
                "  SUM(vw_FILE6_20.QUANT) AS t_q, " & _
                "  SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET, " & _
                "  SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
                "  SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft " & _
                " FROM FILE1_10 INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM " & _
                " left JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE " & _
                " LEFT JOIN FILE1_50 ON FILE1_50.CODE = FILE1_10.[GROUP]" & _
                " inner join file4_10 on file4_10.code = file1_10.code" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"
    ElseIf SEC(1).Value <> 0 Then
        cString = " SELECT FILE1_10.[GROUP]," & _
                " FILE1_50.DESCA ," & _
                " FILE0_81.DESCA ,   " & _
                "  SUM(vw_FILE6_20.QUANT) AS t_q, " & _
                "  SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET, " & _
                "  SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
                "  SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft " & _
                " FROM FILE1_10 INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM" & _
                " left JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX" & _
                " LEFT JOIN FILE1_50 ON FILE1_50.CODE = FILE1_10.[GROUP]" & _
                " inner join file4_10 on file4_10.code = file1_10.code" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"
    ElseIf SEC(2).Value <> 0 Then
        cString = "SELECT FILE1_10.[GROUP]," & _
                " FILE1_50.DESCA," & _
                " File0_82.DESCA ,   " & _
                "  SUM(vw_FILE6_20.QUANT) AS t_q, " & _
                "  SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET, " & _
                "  SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
                "  SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft " & _
                " FROM FILE1_10 INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM" & _
                " left JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE" & _
                " LEFT JOIN FILE1_50 ON FILE1_50.CODE = FILE1_10.[GROUP]" & _
                " inner join file4_10 on file4_10.code = file1_10.code" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"
    ElseIf SEC(3).Value <> 0 Then
        cString = " SELECT FILE1_10.[GROUP]," & _
                " FILE1_50.DESCA, NULL AS mosm ,   " & _
                "  SUM(vw_FILE6_20.QUANT) AS t_q, " & _
                "  SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET, " & _
                "  SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
                "  SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft " & _
                " FROM FILE1_10 INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM" & _
                " INNER JOIN FILE1_50 ON FILE1_50.CODE = FILE1_10.[GROUP]" & _
                " LEFT JOIN File0_82  ON File0_82.CODE = FILE1_10.MODELAGE" & _
                " inner join file4_10 on file4_10.code = file1_10.code" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"
    End If
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    
    If xMosm.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    
    If xMosm.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xmosm22.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
    If xsex.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xage.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If XDOC_DISC.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.DOC_DISC = " & MyParn(XDOC_DISC.BoundText)
    If XOutlet.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " file1_10.Outlet = 1 "
    If xSupp.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGrCust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
    If xGroup.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If xdoc_no.text <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.doc_no = " & MyParn(xdoc_no.text)
    If XSTORE.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " STORE = " & MyParn(XSTORE.BoundText)
    If xSection.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[Section] = " & MyParn(xSection.BoundText)
    If IsDate(xDate1.text) Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[DATE] <= " & DateSq(xdate2.text)
    If xall(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.ISBRANCH = 0 "
    If xonest(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 0 "
    If xonest(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 1 "
    If xcust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.CUST = " & MyParn(xcust.BoundText)
    If XISONLINE.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.online <> 0 "
    If XSHIP.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[SHIP] = " & MyParn(XSHIP.BoundText)
    

    If cWhere <> "" Then cString = cString & " WHERE " & cWhere

    If SEC(3).Value <> 0 Then
        cString = cString & " GROUP BY FILE1_10.[GROUP] , FILE1_50.DESCA ORDER BY FILE1_10.[group] "
    End If
    If SEC(2).Value <> 0 Then
        cString = cString & " GROUP BY FILE1_10.[GROUP] , FILE1_50.DESCA  , File0_82.[DESCA]    ORDER BY FILE1_10.[group]  ,File0_82.[DESCA] "
    End If
    If SEC(1).Value <> 0 Then
        cString = cString & " GROUP BY FILE1_10.[GROUP] , FILE1_50.DESCA  , FILE0_81.[DESCA]    ORDER BY FILE1_10.[group] , FILE0_81.[DESCA]"
    End If
    If SEC(0).Value <> 0 Then
        cString = cString & " GROUP BY FILE1_10.[GROUP] , FILE1_50.DESCA  , FILE1_10.[mosm]    ORDER BY FILE1_10.[group] "
    End If
    

    Set DATA24.Recordset = mycmd(cString, con)
    XTEXT1.text = cString
    
    FIXGRID4
End Sub
Sub FIXGRID4()
With grid4
    .RowHeight(0) = 600
    .WordWrap = True
    .ColHidden(2) = False
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·„Ã„Ê⁄… "
    If SEC(0).Value <> 0 Then .TextMatrix(0, 2) = "„Ê”„ "
    If SEC(2).Value <> 0 Then .TextMatrix(0, 2) = "›∆… ⁄„—Ì…"
    If SEC(1).Value <> 0 Then .TextMatrix(0, 2) = "«·‰Ê⁄"
    If SEC(3).Value <> 0 Then .ColHidden(2) = True
    .TextMatrix(0, 3) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 4) = "ﬁÌ„… „»Ì⁄« "
    .TextMatrix(0, 5) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 6) = "—»Õ „»Ì⁄« "
    .ColWidth(0) = 500
    .ColWidth(1) = 2500
    .ColWidth(2) = 1000
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    .ColHidden(5) = Not bOpt10
    .ColHidden(6) = Not bOpt10
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    For nRow = 1 To .Rows - 1
        .TextMatrix(nRow, 6) = Round(Val(.TextMatrix(nRow, 4)) - Val(.TextMatrix(nRow, 5)), 2)
    Next nRow
    For nCol = 3 To .Cols - 1
        .ColFormat(nCol) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        For nRow = 1 To .Rows - 1
            '.TextMatrix(nRow, nCol) = Format(.TextMatrix(nRow, nCol), "#0.00")
        Next nRow
    Next nCol
    .SubtotalPosition = flexSTAbove
    For i = 3 To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next i
    If .Rows > 1 Then .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub
Private Sub MYLOAD3()
Dim i As Double
Dim cString  As String, cWhere As String
'                           0               1                 2                3               4                5
    If SEC(0).Value Then
        cString = " SELECT FILE1_10.ModelAge" & _
                ", FILE0_82.DESCA" & _
                ", FILE1_10.MOSM" & _
                ", SUM(vw_FILE6_20.QUANT) AS t_q " & _
                ", SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET " & _
                ", SUM(vw_FILE6_20.ITEM_COST) AS t_cost " & _
                ", SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft " & _
                " FROM FILE1_10 " & _
                " INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM" & _
                " LEFT JOIN FILE1_10SC ON FILE1_10SC.CODE = FILE1_10.[SECTION]" & _
                " LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE" & _
                " inner join file4_10 on file4_10.code = file1_10.code" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"
    ElseIf SEC(1).Value Then
        cString = " SELECT FILE1_10.ModelAge" & _
                ", FILE0_82.DESCA" & _
                ", FILE0_81.DESCA   " & _
                ", SUM(vw_FILE6_20.QUANT) AS t_q " & _
                ", SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET " & _
                ", SUM(vw_FILE6_20.ITEM_COST) AS t_cost " & _
                ", SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft " & _
                " FROM FILE1_10 " & _
                " INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM" & _
                " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX" & _
                " LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE" & _
                " inner join file4_10 on file4_10.code = file1_10.code" & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"

    ElseIf SEC(2).Value Or SEC(3).Value Then
        cString = "SELECT FILE1_10.ModelAge" & _
                  ", FILE0_82.DESCA" & _
                  ", NULL AS mosm  " & _
                  ", SUM(vw_FILE6_20.QUANT) AS t_q " & _
                  ", SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET " & _
                  ", SUM(vw_FILE6_20.ITEM_COST) AS t_cost " & _
                  ", SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft " & _
                  " FROM FILE1_10 " & _
                  " INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM " & _
                  " LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE" & _
                  " inner join file4_10 on file4_10.code = file1_10.code" & _
                  " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"
    End If
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    
    If xcust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.CUST = " & MyParn(xcust.BoundText)
    If xMosm.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    If xmosm22.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
    If xsex.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xage.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If XDOC_DISC.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.DOC_DISC = " & MyParn(XDOC_DISC.BoundText)
    If XOutlet.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " file1_10.Outlet = 1 "
    If xSupp.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGrCust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
    If xGroup.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If xdoc_no.text <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.doc_no = " & MyParn(xdoc_no.text)
    If XSTORE.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " STORE = " & MyParn(XSTORE.BoundText)
    If xSection.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[Section] = " & MyParn(xSection.BoundText)
    If IsDate(xDate1.text) Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[DATE] <= " & DateSq(xdate2.text)
    If xall(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.ISBRANCH = 0 "
    If xonest(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 0 "
    If xonest(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 1 "
    If XISONLINE.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.online <> 0 "
    If XSHIP.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[SHIP] = " & MyParn(XSHIP.BoundText)
    

    If cWhere <> "" Then cString = cString & " WHERE " & cWhere

    If SEC(0).Value <> 0 Then
        cString = cString & " GROUP BY FILE1_10.ModelAge , FILE0_82.DESCA , FILE1_10.[MOSM] ORDER BY FILE1_10.ModelAge  "
    End If
    If SEC(1).Value <> 0 Then
        cString = cString & " GROUP BY FILE1_10.ModelAge , FILE0_82.DESCA  , FILE0_81.[DESCA]    ORDER BY FILE1_10.ModelAge "
    End If
    If SEC(3).Value <> 0 Or SEC(2).Value <> 0 Then
        cString = cString & " GROUP BY FILE1_10.ModelAge , FILE0_82.DESCA ORDER BY FILE1_10.ModelAge "
    End If
    
    
    Set DATA23.Recordset = mycmd(cString, con)
    XTEXT1.text = cString
    
    FixGrid3
End Sub
Sub FixGrid3()
With grid3
    .RowHeight(0) = 600
    .WordWrap = True
    
    .ColHidden(3) = False
    .ColHidden(2) = False
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "›∆… ⁄„—Ì…"
    
    If SEC(0).Value <> 0 Then .TextMatrix(0, 2) = "„Ê”„"
    If SEC(1).Value <> 0 Then .TextMatrix(0, 2) = "«·‰Ê⁄"
    If SEC(3).Value <> 0 Then .ColHidden(2) = True
    If SEC(2).Value <> 0 Then .ColHidden(2) = True
    
    .TextMatrix(0, 3) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 4) = "ﬁÌ„… „»Ì⁄« "
    
    .TextMatrix(0, 5) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 6) = "—»Õ „»Ì⁄« "
    
    
    .ColWidth(0) = 500
    .ColWidth(1) = 2500
    .ColWidth(2) = 1000
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    
    .ColHidden(5) = Not bOpt10
    .ColHidden(6) = Not bOpt10
    
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    For nRow = 1 To .Rows - 1
        '.TextMatrix(nRow, 6) = Round(Val(.TextMatrix(nRow, 4)) - Val(.TextMatrix(nRow, 5)), 2)
    Next nRow
    For nCol = 3 To .Cols - 1
        .ColFormat(nCol) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol) = Format(.TextMatrix(nRow, nCol), "#0.00")
        Next nRow
    Next nCol
    .SubtotalPosition = flexSTAbove
    For i = 3 To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next i
    If grid1.Rows > 1 Then .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub
Private Sub MYLOAD2()
Dim i As Double
Dim cString  As String, cWhere As String
'                           0               1                 2                3               4                5
    cString = " SELECT vw_FILE6_20.date " & _
                ", NULL as n2" & _
                ",  COUNT(Distinct vw_FILE6_20.doc_no) " & _
                ", SUM(vw_FILE6_20.QUANT) AS t_q " & _
                ", SUM(vw_FILE6_20.TOTAL_C) AS t_item " & _
                ", SUM(vw_FILE6_20.ITEM_DISCOUNT_C) AS t_disc1 " & _
                ", SUM(vw_FILE6_20.ITEM_DISCOUNT) AS t_dist2 " & _
                ", SUM(vw_FILE6_20.ITEM_DISCOUNT) + SUM(vw_FILE6_20.ITEM_DISCOUNT_C) as n1" & _
                ", SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET " & _
                ", SUM(vw_FILE6_20.ITEM_COST) AS t_cost " & _
                ", SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft " & _
                ", CASE WHEN SUM(vw_FILE6_20.TOTAL_C) <> 0 THEN ROUND((SUM(vw_FILE6_20.ITEM_DISCOUNT) + SUM(vw_FILE6_20.ITEM_DISCOUNT_C)) / SUM(vw_FILE6_20.TOTAL_C),4) ELSE 0 END  * 100 AS DIS_RATE" & _
                ", CASE WHEN SUM(vw_FILE6_20.TOTAL_NET) <> 0  THEN ROUND(SUM(vw_FILE6_20.ITEM_PROFIT)  / SUM(vw_FILE6_20.TOTAL_NET),4) ELSE 0 END * 100 AS PROFIT_RATE " & _
                " FROM  FILE1_10 " & _
                " INNER JOIN vw_FILE6_20 ON FILE1_10.ITEM = vw_FILE6_20.ITEM " & _
                " inner join file4_10 on file4_10.code = file1_10.code " & _
                " INNER JOIN STORE_BR ON vw_FILE6_20.STORE = STORE_BR.CODE"
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    
    If xcust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.CUST = " & MyParn(xcust.BoundText)
    If xMosm.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    If xmosm22.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
    If xsex.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xage.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If XDOC_DISC.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.DOC_DISC = " & MyParn(XDOC_DISC.BoundText)
    If XOutlet.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " file1_10.Outlet = 1 "
    If xSupp.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGrCust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
    If xGroup.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If xdoc_no.text <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.doc_no = " & MyParn(xdoc_no.text)
    If XSTORE.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " STORE = " & MyParn(XSTORE.BoundText)
    If xSection.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[Section] = " & MyParn(xSection.BoundText)
    If IsDate(xDate1.text) Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[DATE] <= " & DateSq(xdate2.text)
    If xall(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.ISBRANCH = 0 "
    If xonest(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 0 "
    If xonest(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 1 "
    If XISONLINE.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.online <> 0 "
    If XSHIP.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[SHIP] = " & MyParn(XSHIP.BoundText)
    
        
    If cWhere <> "" Then cString = cString & " WHERE " & cWhere

    cString = cString & " GROUP BY  vw_FILE6_20.date ORDER BY vw_FILE6_20.date "
    
    Set DATA22.Recordset = mycmd(cString, con)
    XTEXT1.text = cString
    
    FixGrid2

End Sub
Sub FixGrid2()
With GRID2
    .RowHeight(0) = 600
    .WordWrap = True
    .Cols = 13
    
    .TextMatrix(0, 0) = "«· «—ÌÕ"
    .TextMatrix(0, 1) = "«·ÌÊ„"
    .TextMatrix(0, 2) = "⁄œœ »Ê‰« "
    .TextMatrix(0, 3) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 4) = "ﬁÌ„… «·√’‰«›"
    .TextMatrix(0, 5) = "Œ’„ √Êﬂ«“ÊÌ‰"
    .TextMatrix(0, 6) = "Œ’„ »Ê‰« "
    .TextMatrix(0, 7) = "≈Ã„«·Ï ﬁÌ„… «·Œ’„"
    .TextMatrix(0, 8) = "’«›Ï ﬁÌ„… „»Ì⁄«  «·ÌÊ„"
    
    .TextMatrix(0, 9) = " ﬂ·›… «·„»Ì⁄« "
    .TextMatrix(0, 10) = "—»Õ „»Ì⁄« "
    .TextMatrix(0, 11) = "‰”»… «·Œ’„"
    .TextMatrix(0, 12) = "‰”»… «·—»Õ"
    
    .ColHidden(9) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    .ColHidden(12) = Not bOpt10
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1000
    .ColWidth(2) = 800
    .ColWidth(3) = 1000
    .ColWidth(4) = 1300
    .ColWidth(5) = 1300
    .ColWidth(6) = 1300
    .ColWidth(7) = 1300
    .ColWidth(8) = 1300
    .ColWidth(9) = 1300
    .ColWidth(10) = 1300
    .ColWidth(11) = 1100
    .ColWidth(12) = 1100
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .ColDataType(0) = flexDTDate
    For nRow = 1 To .Rows - 1
'        .TextMatrix(nRow, 10) = Round(Val(.TextMatrix(nRow, 8)) - Val(.TextMatrix(nRow, 9)), 2)
'        .TextMatrix(nRow, 7) = Round(Val(.TextMatrix(nRow, 5)) + Val(.TextMatrix(nRow, 6)), 2)
'        If Val(.TextMatrix(nRow, 4)) > 0 Then .TextMatrix(nRow, 11) = Round(Val(.TextMatrix(nRow, 7)) / Val(.TextMatrix(nRow, 4)) * 100, 2)
'        If Val(.TextMatrix(nRow, 4)) > 0 Then .TextMatrix(nRow, 12) = Round(Val(.TextMatrix(nRow, 10)) / Val(.TextMatrix(nRow, 4)) * 100, 2)
    Next nRow
    For nCol = 3 To .Cols - 1
        .ColFormat(nCol) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol) = Format(.TextMatrix(nRow, nCol), "#0.00")
        Next nRow
    Next nCol
    
    .SubtotalPosition = flexSTAbove
    For i = 2 To .Cols - 3
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next i
    
    For nRow = 2 To .Rows - 1
        .TextMatrix(nRow, 1) = ArabicDay(.TextMatrix(nRow, 0))
    Next nRow
    
    If .Rows > 2 Then
        nRow = 1
        '.TextMatrix(nRow, 7) = Round(Val(.TextMatrix(nRow, 5)) + Val(.TextMatrix(nRow, 6)), 2)
        If Val(.TextMatrix(nRow, 4)) > 0 Then .TextMatrix(nRow, 11) = Round(Val(.TextMatrix(nRow, 7)) / Val(.TextMatrix(nRow, 4)) * 100, 2)
        If Val(.TextMatrix(nRow, 8)) > 0 Then .TextMatrix(nRow, 12) = Round(Val(.TextMatrix(nRow, 10)) / Val(.TextMatrix(nRow, 8)) * 100, 2)
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
    End With
End Sub
Private Sub myload0()
Dim i As Double
Dim cString  As String
   
CW = ""
Dim cWhere As String
If IsDate(xDate1.text) Then CW = " AND DATE >= " & DateSq(xDate1.text)
If IsDate(xdate2.text) Then CW = CW & " AND DATE <= " & DateSq(xdate2.text)
    'cF1 = " - Coalesce(( SELECT SUM(VALUE) FROM FILE8_00 INNER JOIN FILE8_00H ON FILE8_00.DOC_NO = FILE8_00H.DOC_NO WHERE STORE = STORE_BR.CODE " & CW & " ),0) "
    cString = " SELECT 1 AS FLAG ," & _
                "STORE_BR.S_BRANCH ," & _
                " STORE_BR.DESCA  , " & _
                " count(Distinct vw_FILE6_20.doc_no) " & _
                " , SUM(vw_FILE6_20.QUANT) AS t_q " & _
                " , SUM(vw_FILE6_20.TOTAL_C) AS t_item " & _
                " , SUM(vw_FILE6_20.ITEM_DISCOUNT_C) AS t_disc1 " & _
                " , SUM(vw_FILE6_20.ITEM_DISCOUNT) AS t_dist2 " & _
                " , SUM(vw_FILE6_20.ITEM_DISCOUNT) + SUM(vw_FILE6_20.ITEM_DISCOUNT_C) as n1" & _
                " , SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET " & _
                " , SUM(vw_FILE6_20.ITEM_COST) AS t_cost " & _
                " , SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft " & _
                " , CASE WHEN SUM(vw_FILE6_20.TOTAL_C) <> 0 THEN ROUND((SUM(vw_FILE6_20.ITEM_DISCOUNT) + SUM(vw_FILE6_20.ITEM_DISCOUNT_C)) / SUM(vw_FILE6_20.TOTAL_C),2) ELSE 0 END  * 100 AS DIS_RATE" & _
                " , CASE WHEN SUM(vw_FILE6_20.TOTAL_NET) <> 0  THEN ROUND(SUM(vw_FILE6_20.ITEM_PROFIT)  / SUM(vw_FILE6_20.TOTAL_NET),2) ELSE 0 END * 100 AS PROFIT_RATE " & _
                " FROM  STORE_BR " & _
                " INNER JOIN vw_FILE6_20 ON STORE_BR.CODE = vw_FILE6_20.STORE" & _
                " INNER JOIN FILE1_10 ON vw_FILE6_20.ITEM = FILE1_10.ITEM" & _
                " INNER JOIN FILE4_10 ON FILE4_10.CODE = FILE1_10.CODE"
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    
    If xcust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.CUST = " & MyParn(xcust.BoundText)
    If xMosm.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    If xmosm22.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
    If xsex.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xage.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If XDOC_DISC.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.DOC_DISC = " & MyParn(XDOC_DISC.BoundText)
    If XOutlet.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " file1_10.Outlet = 1 "
    If xSupp.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGrCust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
    If xGroup.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If xdoc_no.text <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.doc_no = " & MyParn(xdoc_no.text)
    If XSTORE.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " STORE = " & MyParn(XSTORE.BoundText)
    If xSection.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[Section] = " & MyParn(xSection.BoundText)
    If IsDate(xDate1.text) Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[DATE] <= " & DateSq(xdate2.text)
    If xall(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.ISBRANCH = 0 "
    If xonest(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 0 "
    If xonest(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 1 "
    If XSHIP.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[SHIP] = " & MyParn(XSHIP.BoundText)
    
    
    If XISONLINE.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.online <> 0 "
    If XSHIP.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[SHIP] = " & MyParn(XSHIP.BoundText)
    
        
    If cWhere <> "" Then cString = cString & " WHERE " & cWhere
    
    cString = cString & " GROUP BY  STORE_BR.S_BRANCH ,STORE_BR.CODE , STORE_BR.DESCA ORDER BY STORE_BR.S_BRANCH "
    Set data20.Recordset = mycmd(cString, con)
    XTEXT1.text = cString
    
    FIXGRID0
End Sub
Sub FIXGRID0()
With GRID0
    .RowHeight(0) = 600
    .WordWrap = True
    .ColHidden(0) = True
    .TextMatrix(0, 1) = "ﬂÊœ"
    .TextMatrix(0, 2) = "«·›—⁄ "
    
    
    .TextMatrix(0, 3) = "⁄œœ »Ê‰« "
    .TextMatrix(0, 4) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 5) = "ﬁÌ„… «·√’‰«›"
    .TextMatrix(0, 6) = "Œ’„ «’‰«›"
    .TextMatrix(0, 7) = "Œ’„ »Ê‰« "
    .TextMatrix(0, 8) = "≈Ã„«·Ï ﬁÌ„… «·Œ’„"
    .TextMatrix(0, 9) = "’«›Ï ﬁÌ„… „»Ì⁄«  "
    
    .TextMatrix(0, 10) = " ﬂ·›… «·„»Ì⁄« "
    .TextMatrix(0, 11) = "—»Õ „»Ì⁄« "
    .TextMatrix(0, 12) = "‰”»… «·Œ’„"
    .TextMatrix(0, 13) = "‰”»… «·—»Õ"
    
    .ColWidth(1) = 600
    .ColWidth(2) = 2000
    .ColWidth(3) = 800
    .ColWidth(4) = 1000
    .ColWidth(5) = 1300
    .ColWidth(6) = 1300
    .ColWidth(7) = 1300
    .ColWidth(8) = 1300
    .ColWidth(9) = 1300
    .ColWidth(10) = 1300
    .ColWidth(11) = 1300
    .ColWidth(12) = 900
    .ColWidth(13) = 900
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
'    .ColDataType(0) = flexDTDate
    For nRow = 1 To .Rows - 1
'        .TextMatrix(nRow, 11) = Round(Val(.TextMatrix(nRow, 9)) - Val(.TextMatrix(nRow, 10)), 2)
'        .TextMatrix(nRow, 8) = Round(Val(.TextMatrix(nRow, 6)) + Val(.TextMatrix(nRow, 7)), 2)
'        If Val(.TextMatrix(nRow, 5)) > 0 Then .TextMatrix(nRow, 12) = Round(Val(.TextMatrix(nRow, 8)) / Val(.TextMatrix(nRow, 5)) * 100, 0)
'        If Val(.TextMatrix(nRow, 5)) > 0 Then .TextMatrix(nRow, 13) = Round(Val(.TextMatrix(nRow, 11)) / Val(.TextMatrix(nRow, 5)) * 100, 0)
    Next nRow
    
    For nCol = 3 To .Cols - 3
        .ColFormat(nCol) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
    Next nCol
    
    .ColHidden(10) = Not bOpt10
    .ColHidden(11) = Not bOpt10
    .ColHidden(13) = Not bOpt10
    
    .SubtotalPosition = flexSTAbove
    For i = 3 To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
'        .Subtotal flexSTSum, 0, i, "#0", vbGreen, , True, "  "
    Next i
    
    If .Rows > 1 Then
        .TextMatrix(1, 8) = Round(Val(.TextMatrix(1, 6)) + Val(.TextMatrix(1, 7)), 2)
        If Val(.TextMatrix(1, 5)) <> 0 Then .TextMatrix(1, 12) = Round(Val(.TextMatrix(1, 8)) / Val(.TextMatrix(1, 5)) * 100, 2)
        If Val(.TextMatrix(1, 9)) <> 0 Then .TextMatrix(1, 13) = Round(Val(.TextMatrix(1, 11)) / Val(.TextMatrix(1, 9)) * 100, 2)
        .Cell(flexcpAlignment, 1, 0, 1, .Cols - 1) = 7
    End If
    End With
End Sub
Private Sub MYLOAD1()
Dim i As Double
Dim cString  As String, cWhere As String
'                           0                           1                 3
    cString = " SELECT YEAR(vw_FILE6_20.date)" & _
                ",  MONTH(vw_FILE6_20.date)" & _
                ", COUNT(Distinct vw_FILE6_20.doc_no) " & _
                ", SUM(vw_FILE6_20.QUANT) AS t_q " & _
                ", SUM(vw_FILE6_20.TOTAL_C) AS t_item " & _
                ", SUM(vw_FILE6_20.ITEM_DISCOUNT_C) AS t_disc1 " & _
                ", SUM(vw_FILE6_20.ITEM_DISCOUNT) AS t_dist2 " & _
                ", SUM(vw_FILE6_20.ITEM_DISCOUNT) + SUM(vw_FILE6_20.ITEM_DISCOUNT_C) as n1" & _
                ", SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET " & _
                ", SUM(vw_FILE6_20.ITEM_COST) AS t_cost " & _
                ", SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft " & _
                ", CASE WHEN SUM(vw_FILE6_20.TOTAL_C) <> 0 THEN ROUND((SUM(vw_FILE6_20.ITEM_DISCOUNT) + SUM(vw_FILE6_20.ITEM_DISCOUNT_C)) / SUM(vw_FILE6_20.TOTAL_C),4) ELSE 0 END  * 100 AS DIS_RATE" & _
                ", CASE WHEN SUM(vw_FILE6_20.TOTAL_NET) <> 0  THEN ROUND(SUM(vw_FILE6_20.ITEM_PROFIT)  / SUM(vw_FILE6_20.TOTAL_NET),4) ELSE 0 END * 100 AS PROFIT_RATE " & _
                " FROM  STORE_BR " & _
                " INNER JOIN vw_FILE6_20 ON STORE_BR.CODE = vw_FILE6_20.STORE" & _
                " INNER JOIN FILE1_10 ON vw_FILE6_20.ITEM = FILE1_10.ITEM" & _
                " INNER JOIN FILE4_10 ON FILE4_10.CODE = FILE1_10.CODE"
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    
    If xcust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.CUST = " & MyParn(xcust.BoundText)
    If xMosm.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    If xmosm22.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
    If xsex.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xage.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If XDOC_DISC.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.DOC_DISC = " & MyParn(XDOC_DISC.BoundText)
    If XOutlet.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " file1_10.Outlet = 1 "
    If xSupp.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGrCust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
    If xGroup.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If xdoc_no.text <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.doc_no = " & MyParn(xdoc_no.text)
    If XSTORE.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " STORE = " & MyParn(XSTORE.BoundText)
    If xSection.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[Section] = " & MyParn(xSection.BoundText)
    If IsDate(xDate1.text) Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[DATE] <= " & DateSq(xdate2.text)
    If xall(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.ISBRANCH = 0 "
    If xonest(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 0 "
    If xonest(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 1 "
    If XISONLINE.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.online <> 0 "
    If XSHIP.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " vw_FILE6_20.[SHIP] = " & MyParn(XSHIP.BoundText)
    
    

    If cWhere <> "" Then cString = cString & " WHERE " & cWhere

    cString = cString & " GROUP BY  YEAR(vw_FILE6_20.date ), MONTH (vw_FILE6_20.date )  ORDER BY YEAR(vw_FILE6_20.date ), MONTH (vw_FILE6_20.date )  "
    Set DATA21.Recordset = myRecordSet(cString, con)
    XTEXT1.text = cString
    FixGrid1

End Sub
Sub FixGrid1()
With grid1
    .RowHeight(0) = 600
    .WordWrap = True
    .Cols = 13
    .TextMatrix(0, 0) = "«·”‰…"
    .TextMatrix(0, 1) = "«·‘Â—"
    .TextMatrix(0, 2) = "⁄œœ »Ê‰« "
    .TextMatrix(0, 3) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 4) = "ﬁÌ„… «·√’‰«›"
    .TextMatrix(0, 5) = "Œ’„ √Êﬂ«“ÊÌ‰"
    .TextMatrix(0, 6) = "Œ’„ »Ê‰« "
    .TextMatrix(0, 7) = "≈Ã„«·Ï ﬁÌ„… «·Œ’„"
    .TextMatrix(0, 8) = "’«›Ï ﬁÌ„… „»Ì⁄«  «·ÌÊ„"
    
    .TextMatrix(0, 9) = " ﬂ·›… «·„»Ì⁄« "
    .TextMatrix(0, 10) = "—»Õ „»Ì⁄« "
    .TextMatrix(0, 11) = "‰”»… «·Œ’„"
    .TextMatrix(0, 12) = "‰”»… «·—»Õ"
    
    .ColHidden(9) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    .ColHidden(12) = Not bOpt10
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1000
    .ColWidth(2) = 800
    .ColWidth(3) = 1000
    .ColWidth(4) = 1300
    .ColWidth(5) = 1300
    .ColWidth(6) = 1300
    .ColWidth(7) = 1300
    .ColWidth(8) = 1300
    .ColWidth(9) = 1300
    .ColWidth(10) = 1300
    .ColWidth(11) = 1100
    .ColWidth(12) = 1100
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .ColDataType(0) = flexDTDate
    
    For nRow = 1 To .Rows - 1
'        .TextMatrix(nRow, 10) = Round(Val(.TextMatrix(nRow, 8)) - Val(.TextMatrix(nRow, 9)), 2)
'        .TextMatrix(nRow, 7) = Round(Val(.TextMatrix(nRow, 5)) + Val(.TextMatrix(nRow, 6)), 2)
'        If Val(.TextMatrix(nRow, 4)) > 0 Then .TextMatrix(nRow, 11) = Round(Val(.TextMatrix(nRow, 7)) / Val(.TextMatrix(nRow, 4)) * 100, 2)
'        If Val(.TextMatrix(nRow, 4)) > 0 Then .TextMatrix(nRow, 12) = Round(Val(.TextMatrix(nRow, 10)) / Val(.TextMatrix(nRow, 4)) * 100, 2)
    Next nRow
    
    For nCol = 3 To .Cols - 1
        .ColFormat(nCol) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
    Next nCol
    
    .SubtotalPosition = flexSTAbove
    For i = 2 To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0.00", vbRed, vbYellow, True, "  "
    Next i
    If .Rows > 1 Then
        '.TextMatrix(.Rows - 1, 7) = Round(Val(.TextMatrix(.Rows - 1, 5)) + Val(.TextMatrix(.Rows - 1, 6)), 2)
        If Val(.TextMatrix(1, 4)) > 0 Then .TextMatrix(1, 11) = Round((Val(.TextMatrix(1, 7)) / Val(.TextMatrix(1, 4))) * 100, 2)
        If Val(.TextMatrix(1, 8)) > 0 Then .TextMatrix(1, 12) = Round((Val(.TextMatrix(1, 10)) / Val(.TextMatrix(1, 8))) * 100, 2)
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
    End With
End Sub

Private Sub XTEXT1_GotFocus()
    XTEXT1.SelStart = 0
    XTEXT1.SelLength = Len(XTEXT1.text)

End Sub
Private Sub MYLOAD11()
Dim i As Double
    Dim cString  As String, cWhere As String '                           0               1                 2                3               4                5
    cString = " SELECT VW_FILE6_20.date," & _
                " STORE_ALL.desca," & _
                " COUNT(Distinct vw_FILE6_20.doc_no), " & _
                "  SUM(vw_FILE6_20.QUANT) AS t_q, " & _
                "  SUM(vw_FILE6_20.TOTAL_C) AS t_item, " & _
                "  SUM(vw_FILE6_20.ITEM_DISCOUNT_C) AS t_disc1, " & _
                "  SUM(vw_FILE6_20.ITEM_DISCOUNT) AS t_dist2, " & _
                "  SUM(vw_FILE6_20.ITEM_DISCOUNT) + SUM(vw_FILE6_20.ITEM_DISCOUNT_C) as n1," & _
                "  SUM(vw_FILE6_20.TOTAL_NET) AS TOTAL_NET, " & _
                "  SUM(vw_FILE6_20.ITEM_COST) AS t_cost, " & _
                "  SUM(vw_FILE6_20.ITEM_PROFIT) AS t_proft, " & _
                "  CASE WHEN SUM(vw_FILE6_20.TOTAL_C) <> 0 THEN ROUND((SUM(vw_FILE6_20.ITEM_DISCOUNT) + SUM(vw_FILE6_20.ITEM_DISCOUNT_C)) / SUM(vw_FILE6_20.TOTAL_C),4) ELSE 0 END  * 100 AS DIS_RATE," & _
                "  CASE WHEN SUM(vw_FILE6_20.TOTAL_NET) <> 0  THEN ROUND(SUM(vw_FILE6_20.ITEM_PROFIT)  / SUM(vw_FILE6_20.TOTAL_NET),4) ELSE 0 END * 100 AS PROFIT_RATE " & _
                " FROM  FILE1_10 " & _
                " INNER JOIN VW_FILE6_20 ON FILE1_10.ITEM = VW_FILE6_20.ITEM" & _
                " inner join file4_10 on file4_10.code = file1_10.code" & _
                " inner join STORE_ALL on STORE_ALL.code = VW_FILE6_20.store" & _
                " INNER JOIN STORE_BR ON VW_FILE6_20.STORE = STORE_BR.CODE"
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    If xcust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.CUST = " & MyParn(xcust.BoundText)
    If xMosm.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    If xmosm22.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
    If xsex.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xage.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If XDOC_DISC.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.DOC_DISC = " & MyParn(XDOC_DISC.BoundText)
    If XOutlet.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " file1_10.Outlet = 1 "
    If xSupp.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGrCust.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
    If xGroup.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If xdoc_no.text <> "" Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.doc_no = " & MyParn(xdoc_no.text)
    If XSTORE.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " STORE = " & MyParn(XSTORE.BoundText)
    If xSection.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[Section] = " & MyParn(xSection.BoundText)
    If IsDate(xDate1.text) Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.[DATE] <= " & DateSq(xdate2.text)
    If xall(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.ISBRANCH = 0 "
    If xonest(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 0 "
    If xonest(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.ISONEST = 1 "
    If XISONLINE.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.online <> 0 "
    If XSHIP.BoundText <> "" Then cWhere = cWhere & Tr(cWhere) & " VW_FILE6_20.[SHIP] = " & MyParn(XSHIP.BoundText)
    

    If cWhere <> "" Then cString = cString & " WHERE " & cWhere

    cString = cString & " GROUP BY  VW_FILE6_20.date , STORE_ALL.code, STORE_ALL.desca" & _
                        " ORDER BY VW_FILE6_20.date , STORE_ALL.code "
    Set data31.Recordset = mycmd(cString, con)
    XTEXT1.text = cString
    
    FixGrid11

End Sub
Sub FixGrid11()
With GRID11
    .RowHeight(0) = 600
    .WordWrap = True
    .Cols = 13
    .TextMatrix(0, 0) = "«· «—ÌÕ"
    .TextMatrix(0, 1) = "«·›—⁄ "
    .TextMatrix(0, 2) = "⁄œœ »Ê‰« "
    .TextMatrix(0, 3) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 4) = "ﬁÌ„… «·√’‰«›"
    .TextMatrix(0, 5) = "Œ’„ √Êﬂ«“ÊÌ‰"
    .TextMatrix(0, 6) = "Œ’„ »Ê‰« "
    .TextMatrix(0, 7) = "≈Ã„«·Ï ﬁÌ„… «·Œ’„"
    .TextMatrix(0, 8) = "’«›Ï ﬁÌ„… „»Ì⁄«  «·ÌÊ„"
    
    .TextMatrix(0, 9) = " ﬂ·›… «·„»Ì⁄« "
    .TextMatrix(0, 10) = "—»Õ „»Ì⁄« "
    .TextMatrix(0, 11) = "‰”»… «·Œ’„"
    .TextMatrix(0, 12) = "‰”»… «·—»Õ"
    
    .ColHidden(9) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    .ColHidden(12) = Not bOpt10
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 2000
    .ColWidth(2) = 800
    .ColWidth(3) = 1000
    .ColWidth(4) = 1300
    .ColWidth(5) = 1300
    .ColWidth(6) = 1300
    .ColWidth(7) = 1300
    .ColWidth(8) = 1300
    .ColWidth(9) = 1300
    .ColWidth(10) = 1300
    .ColWidth(11) = 1100
    .ColWidth(12) = 1100
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .ColDataType(0) = flexDTDate
    For nRow = 1 To .Rows - 1
'        .TextMatrix(nRow, 10) = Round(Val(.TextMatrix(nRow, 8)) - Val(.TextMatrix(nRow, 9)), 2)
'        .TextMatrix(nRow, 7) = Round(Val(.TextMatrix(nRow, 5)) + Val(.TextMatrix(nRow, 6)), 2)
'        If Val(.TextMatrix(nRow, 4)) > 0 Then .TextMatrix(nRow, 11) = Round(Val(.TextMatrix(nRow, 7)) / Val(.TextMatrix(nRow, 4)) * 100, 2)
'        If Val(.TextMatrix(nRow, 4)) > 0 Then .TextMatrix(nRow, 12) = Round(Val(.TextMatrix(nRow, 10)) / Val(.TextMatrix(nRow, 4)) * 100, 2)
    Next nRow
    
    For nCol = 3 To .Cols - 1
        .ColFormat(nCol) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
    Next nCol
    
    .SubtotalPosition = flexSTAbove
    For i = 2 To .Cols - 3
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next i
    
    If .Rows > 2 Then
        'nRow = 1
        '.TextMatrix(nRow, 7) = Round(Val(.TextMatrix(nRow, 5)) + Val(.TextMatrix(nRow, 6)), 2)
        If .ValueMatrix(1, 4) > 0 Then .TextMatrix(1, 11) = Round(Val(.TextMatrix(1, 7)) / Val(.TextMatrix(1, 4)) * 100, 2)
        If .ValueMatrix(1, 8) > 0 Then .TextMatrix(1, 12) = Round(Val(.TextMatrix(1, 10)) / Val(.TextMatrix(1, 8)) * 100, 2)
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
    End With
End Sub
