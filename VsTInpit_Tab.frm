VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VsTInpit_Tab 
   Caption         =   " ﬁ«—Ì— «·Ê«—œ „Ã„⁄…"
   ClientHeight    =   10230
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   18585
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
   ScaleHeight     =   10230
   ScaleWidth      =   18585
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2265
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   45
      Width           =   13740
      Begin VB.TextBox XGRMODEL 
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
         TabIndex        =   38
         Top             =   270
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
         Left            =   2880
         MaxLength       =   15
         TabIndex        =   36
         Top             =   270
         Width           =   1545
      End
      Begin VB.TextBox XMOSM2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1710
         Width           =   825
      End
      Begin VB.ComboBox xmosm20 
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
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1710
         Width           =   825
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   7755
         TabIndex        =   2
         Top             =   1395
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   8520
         TabIndex        =   3
         Top             =   1035
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
         Left            =   135
         TabIndex        =   4
         Top             =   630
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   315
         Left            =   8520
         TabIndex        =   5
         Top             =   690
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
         Left            =   135
         TabIndex        =   12
         Top             =   990
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XSECTION 
         Height          =   315
         Left            =   135
         TabIndex        =   14
         Top             =   1350
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker xdate1 
         Height          =   375
         Left            =   10215
         TabIndex        =   21
         Top             =   270
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   661
         _Version        =   393216
         MousePointer    =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   445317121
         CurrentDate     =   43097
      End
      Begin MSComCtl2.DTPicker xdate2 
         Height          =   375
         Left            =   8370
         TabIndex        =   22
         Top             =   270
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   661
         _Version        =   393216
         MousePointer    =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   445317121
         CurrentDate     =   43097
      End
      Begin MSDataListLib.DataCombo xage 
         Height          =   315
         Left            =   10005
         TabIndex        =   28
         Top             =   1755
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
         Left            =   7740
         TabIndex        =   29
         Top             =   1755
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·›—⁄"
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
         Index           =   8
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ „Ã„Ê⁄… "
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
         Left            =   1725
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   315
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ „ÊœÌ·"
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
         Left            =   4515
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   315
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "„Ê”„ œ«Œ·Ï"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1755
         Width           =   1050
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
         Left            =   9315
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1800
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
         Index           =   3
         Left            =   12060
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1800
         Width           =   1005
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
         Index           =   2
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1755
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬁ”„ "
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
         Index           =   1
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1395
         Width           =   600
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·› —… :"
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
         Left            =   12060
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label2 
         Caption         =   "„Ê”„ «·„ÊœÌ·"
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
         Index           =   6
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   720
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
         Left            =   12060
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1065
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
         Left            =   12060
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1440
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
         Left            =   12060
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   675
         Width           =   1545
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   13635
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   2250
      Width           =   4065
      Begin VB.OptionButton SEC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   225
         Value           =   -1  'True
         Width           =   645
      End
      Begin VB.OptionButton SEC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3105
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   225
         Width           =   870
      End
      Begin VB.OptionButton SEC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   225
         Width           =   735
      End
      Begin VB.OptionButton SEC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·›∆… «·⁄„—Ì…"
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
         Index           =   2
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   225
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      Height          =   690
      Left            =   4860
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   2295
      Width           =   2805
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "query "
         Height          =   285
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   270
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
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   270
         Width           =   870
      End
   End
   Begin VB.Frame fmBranch 
      Height          =   690
      Left            =   7695
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   2295
      Width           =   2220
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   225
         Width           =   1950
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2190
      Left            =   90
      TabIndex        =   55
      Top             =   90
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   3863
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "’Ê—… «·„ÊœÌ·"
      TabPicture(0)   =   "VsTInpit_Tab.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ImgX1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "query "
      TabPicture(1)   =   "VsTInpit_Tab.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "textquery"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.TextBox textquery 
         Alignment       =   1  'Right Justify
         Height          =   1770
         Left            =   45
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   360
         Width           =   3660
      End
      Begin ImgXCtrl6.ImgXCtrl ImgX1 
         Height          =   1770
         Left            =   -74910
         TabIndex        =   56
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   3122
         BackColor       =   16777215
         BorderStyle     =   4
         AutoZoom        =   -1  'True
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         AutoZoomType    =   3
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "mrmind"
         LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
      End
   End
   Begin VB.Frame Frame5 
      Height          =   690
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   3960
      Visible         =   0   'False
      Width           =   1635
      Begin MSAdodcLib.Adodc DATA10 
         Height          =   330
         Left            =   -540
         Top             =   585
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
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
         Left            =   -405
         Top             =   495
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
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
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   -315
         Top             =   540
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
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
         Left            =   -180
         Top             =   540
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
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
      Begin MSAdodcLib.Adodc DATA15 
         Height          =   330
         Left            =   0
         Top             =   450
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
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
      Begin MSAdodcLib.Adodc DATA16 
         Height          =   330
         Left            =   -45
         Top             =   450
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
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
      Begin MSAdodcLib.Adodc DATA17 
         Height          =   330
         Left            =   0
         Top             =   540
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
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
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2250
      Width           =   4740
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1080
         Picture         =   "VsTInpit_Tab.frx":0038
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3420
         Picture         =   "VsTInpit_Tab.frx":2823
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "VsTInpit_Tab.frx":4D15
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   135
         Width           =   1005
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2295
         Picture         =   "VsTInpit_Tab.frx":7181
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   135
         Width           =   1095
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   9900
      Width           =   18585
      _ExtentX        =   32782
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   32
      Top             =   9705
      Width           =   18585
      _ExtentX        =   32782
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
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
      Caption         =   "DATA12"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   180
      TabIndex        =   40
      Top             =   3060
      Width           =   17565
      _ExtentX        =   30983
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "„ÊœÌ·« "
      TabPicture(0)   =   "VsTInpit_Tab.frx":95AB
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grid0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "„’«‰⁄"
      TabPicture(1)   =   "VsTInpit_Tab.frx":95C7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GRID1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   " «·√ﬁ”«„"
      TabPicture(2)   =   "VsTInpit_Tab.frx":95E3
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GRID2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "«·„Ã„Ê⁄« "
      TabPicture(3)   =   "VsTInpit_Tab.frx":95FF
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "GRID3"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   " ›∆«  ⁄„—Ì…"
      TabPicture(4)   =   "VsTInpit_Tab.frx":961B
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "GRID4"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   " ›’Ì·Ï „Ã„⁄ - „ÊœÌ·«  - —»Õ"
      TabPicture(5)   =   "VsTInpit_Tab.frx":9637
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "GRID5"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "«·›—Ê⁄"
      TabPicture(6)   =   "VsTInpit_Tab.frx":9653
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "GRID6"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "„Ã„Ê⁄«  „ÊœÌ·« "
      TabPicture(7)   =   "VsTInpit_Tab.frx":966F
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "GRID7"
      Tab(7).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid0 
         Bindings        =   "VsTInpit_Tab.frx":968B
         Height          =   6030
         Left            =   90
         TabIndex        =   41
         Top             =   720
         Width           =   17205
         _cx             =   30348
         _cy             =   10636
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
      Begin VSFlex7Ctl.VSFlexGrid GRID1 
         Bindings        =   "VsTInpit_Tab.frx":969F
         Height          =   6030
         Left            =   -74910
         TabIndex        =   42
         Top             =   720
         Width           =   17205
         _cx             =   30348
         _cy             =   10636
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
      Begin VSFlex7Ctl.VSFlexGrid GRID2 
         Bindings        =   "VsTInpit_Tab.frx":96B3
         Height          =   6030
         Left            =   -74955
         TabIndex        =   43
         Top             =   720
         Width           =   17205
         _cx             =   30348
         _cy             =   10636
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
      Begin VSFlex7Ctl.VSFlexGrid GRID3 
         Bindings        =   "VsTInpit_Tab.frx":96C7
         Height          =   6030
         Left            =   -74910
         TabIndex        =   44
         Top             =   675
         Width           =   17205
         _cx             =   30348
         _cy             =   10636
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
      Begin VSFlex7Ctl.VSFlexGrid GRID4 
         Bindings        =   "VsTInpit_Tab.frx":96DB
         Height          =   6030
         Left            =   -74910
         TabIndex        =   45
         Top             =   720
         Width           =   17205
         _cx             =   30348
         _cy             =   10636
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
      Begin VSFlex7Ctl.VSFlexGrid GRID5 
         Bindings        =   "VsTInpit_Tab.frx":96EF
         Height          =   6075
         Left            =   -74910
         TabIndex        =   46
         Top             =   630
         Width           =   17205
         _cx             =   30348
         _cy             =   10716
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
      Begin VSFlex7Ctl.VSFlexGrid GRID6 
         Bindings        =   "VsTInpit_Tab.frx":9703
         Height          =   6075
         Left            =   -74955
         TabIndex        =   47
         Top             =   675
         Width           =   17205
         _cx             =   30348
         _cy             =   10716
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
      Begin VSFlex7Ctl.VSFlexGrid GRID7 
         Bindings        =   "VsTInpit_Tab.frx":9717
         Height          =   6075
         Left            =   -74910
         TabIndex        =   48
         Top             =   675
         Width           =   17205
         _cx             =   30348
         _cy             =   10716
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
   End
   Begin MSAdodcLib.Adodc DATA8 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
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
      Height          =   735
      Left            =   9945
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2250
      Width           =   3660
      Begin VB.OptionButton xall 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "›—Ê⁄ ›ﬁÿ"
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
         Index           =   0
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   270
         Width           =   1050
      End
      Begin VB.OptionButton xall 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " ÊﬂÌ·«  ›ﬁÿ"
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
         Height          =   225
         Index           =   1
         Left            =   1035
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   270
         Width           =   1230
      End
      Begin VB.OptionButton xall 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   270
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Label xModelNo 
      Alignment       =   1  'Right Justify
      Caption         =   "Label6"
      Height          =   285
      Left            =   1125
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   2310
   End
End
Attribute VB_Name = "VsTInpit_Tab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearchItem As New Search3
Dim oSearch As New Search3
Dim con As New ADODB.Connection
Private Sub CMD_PRINT_Click()
End Sub
Private Sub cmdExel_Click()
    Select Case SSTab1.Tab
        Case 0
            ToFileExel2 GRID0, , , , , 1.1, , , , , , Me
        Case 1
            ToFileExel2 grid1, , , , , 1.1, , , , , , Me
        Case 2
            ToFileExel2 GRID2, , , , , 1.1, , , , , , Me
        Case 3
            ToFileExel2 GRID3, , , , , 1.1, , , , , , Me
        Case 4
            ToFileExel2 GRID4, , , , , 1.1, , , , , , Me
        Case 5
            ToFileExel2 GRID5, , , , , 1.1, , , , , , Me
        Case 6
            ToFileExel2 grid6, , , , , 1.1, , , , , , Me
        Case 7
            ToFileExel2 grid7, , , , , 1.1, , , , , , Me
    End Select
End Sub
Private Sub CmdExit_Click()
    Unload Me
    Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    
    Select Case SSTab1.Tab
        Case 0
            myload0
        Case 1
            MYLOAD1
        Case 2
            MYLOAD2
        Case 3
            MYLOAD3
        Case 4
            MYLOAD4
        Case 5
            myload5
        Case 6
            myload6
        Case 7
            myload7
    End Select
    
    
'    myload0
'    MYLOAD1
'    MYLOAD2
'
'    MYLOAD3
'    myload4
'    myload5
'    myload6
'    myload7
End Sub
Private Sub cmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    Dim cHead3 As String
    cHead1 = Me.Caption & xStore.text & " " & xsex.text & " " & xage.text & "  " & xModelFact.text
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.Value, "DD-MM-YYYY") & " Õ Ï  «—ÌŒ " & Format(xDate2.Value, "DD-MM-YYYY")
    cHead3 = ""
    If xFact.MatchedWithList Then cHead3 = cHead3 & " „’‰⁄ " & xFact.text
    If xSupp.MatchedWithList Then cHead3 = cHead3 & " „Ê—œ " & xSupp.text
    If xGroup.MatchedWithList Then cHead3 = cHead3 & " „Ã„Ê⁄…" & xGroup.text
    If XSECTION.text <> "" Then cHead3 = cHead3 & " ﬁ”„ " & XSECTION.text
    If xMosm.MatchedWithList Then cHead3 = cHead3 & " „Ê”„ " & xMosm.text
    PrintGrdNew.DOPRINT grid1, 0.8, -2, cHead1, cHead2, , , , False, 9
    PrintGrdNew.Show 1
End Sub
Private Sub Form_Load()
    openCon con
    
    textquery.Visible = bSupermode And lSupperVisor
    Check1.Visible = bSupermode And lSupperVisor
    
    xmosm20.AddItem " "
    xmosm20.AddItem "S"
    xmosm20.AddItem "W"
    xmosm20.AddItem "M"
    
    Set data8.Recordset = myRecordSet("Select Code,DescA From File0_82  order by Desca", con)
    Set xage.RowSource = data8
    xage.ListField = "Desca"
    xage.BoundColumn = "Code"
    
    Set DATA6.Recordset = myRecordSet("Select Code,DescA From File0_81  order by Desca", con)
    Set xsex.RowSource = DATA6
    xsex.ListField = "Desca"
    xsex.BoundColumn = "Code"
    
    Set DATA7.Recordset = myRecordSet("Select Code,DescA From STORE_ALL WHERE " & cUserStore, con)
    Set xStore.RowSource = DATA7
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
    If cBranch <> "00" Then
        xStore.BoundText = cBranchStore
        xStore.Enabled = False
    End If
    
    Set data1.Recordset = myRecordSet("Select Code,DescA From File1_10SC ORDER BY code ", con)
    Set XSECTION.RowSource = data1
    XSECTION.ListField = "Desca"
    XSECTION.BoundColumn = "Code"
    
    Set DATA2.Recordset = myRecordSet("Select Code,DescA From File4_10 order by Desca", con)
    Set xSupp.RowSource = DATA2
    xSupp.ListField = "Desca"
    xSupp.BoundColumn = "Code"
    
    Set DATA3.Recordset = myRecordSet("Select Code,DescA From File1_50 ORDER BY DESCA", con)
    Set xGroup.RowSource = DATA3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    Set DATA4.Recordset = myRecordSet("Select mosm ,descA From mosm ORDER BY date DESC ", con)
    Set xMosm.RowSource = DATA4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm
    
    
    Set DATA5.Recordset = myRecordSet("Select code ,desca From fact ORDER BY DESCA ", con)
    Set xFact.RowSource = DATA5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    
    GRID0.Rows = 2
    grid1.Rows = 2
    GRID2.Rows = 2
    GRID3.Rows = 2
    GRID4.Rows = 2
    GRID5.Rows = 2
    grid6.Rows = 2
    grid7.Rows = 2

    Set GRID0.DataSource = data10
    Set grid1.DataSource = DATA11
    Set GRID2.DataSource = data12
    Set GRID3.DataSource = DATA13
    Set GRID4.DataSource = DATA14
    Set GRID5.DataSource = DATA15
    Set grid6.DataSource = DATA16
    Set grid7.DataSource = DATA17

    FIXGRID0
    FixGrid1
    FixGrid2
    FixGrid3
    FIXGRID4
    FIXGRID5
    FIXGRID6
    FIXGRID7
    
    fmBranch.Visible = lSupperVisor
End Sub
Private Sub myload0()
Dim cString  As String
Dim cField As String, cWhere As String, cWhereField As String
With GRID0
    If IsDate(xDate1.Value) Then
        cWhereField = " AND DATE >= " & DateSq(xDate1.Value)
    End If
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(FILE1_11_ALL.[IN] * FILE1_11_ALL.COST ) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST ) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(([IN]-[OUT]) * FILE1_11_ALL.COST ) ")

    cField = cField & ",0," & _
            myiif("(type = '6'   )" & cWhereField, "([OUT]) ")

    cField = cField & "," & _
            myiif("(type = '6'   )" & cWhereField, "([OUT] * FILE1_11_ALL.COST ) ")
        
    cString = " SELECT  FILE4_10.DESCA AS supldesca, FACT.DESCA AS factdesca, FILE1_10.MOSM, FILE1_10.MODELFACT0, FILE1_10.desca " & _
                 cField & " , sum([in]-out) , sum(([in]-out) * FILE1_11_ALL.COST) , sum(([in]-out) * FILE1_10.COSTitem) , FILE1_10.MODELNO " & _
                " FROM  FILE1_10 INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE" & _
                " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
                " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM " & _
                " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                " WHERE file1_10.ISNOITEM = 0 "
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    If xMosm.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING (FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    
    If xmosm20.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm20.text)
            
    If XSECTION.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xStore.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " [store] = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    
    If xage.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If XOutlet.Value = 1 Then cWhere = cWhere & Tr(cWhere) & " Outlet = 1 "


    If IsDate(xDate2.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xDate2.Value)
    
    If xall(0).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 1 "
    ElseIf xall(1).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 0 "
    End If
    
    If cWhere <> "" Then
        cString = cString & " AND " & cWhere
    End If
    cString = cString & " GROUP BY FILE4_10.CODE, FILE4_10.DESCA, FACT.DESCA, FILE1_10.MOSM, FILE1_10.MODELFACT0, FILE1_10.desca , FILE1_10.MODELNO ORDER BY FILE4_10.CODE, FILE4_10.DESCA, FACT.DESCA, FILE1_10.MOSM, FILE1_10.MODELFACT0 "
    Set data10.Recordset = myRecordSet(cString, con)
End With
FIXGRID0
If GRID0.Rows > 0 Then GRID0.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FIXGRID0()
With GRID0
    .Cols = 18
    .ColHidden(17) = True
    .RowHeight(0) = 1000
    .RowHidden(1) = True
    .WordWrap = True
    .ColHidden(1) = True
    .TextMatrix(0, 0) = "«·„Ê—œ"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .TextMatrix(0, 2) = "„Ê”„"
    .TextMatrix(0, 3) = "„ÊœÌ·"
    .TextMatrix(0, 4) = "«·’‰›"
    
    .TextMatrix(0, 5) = "ﬂ„Ì… Ê«—œ"
    .TextMatrix(0, 6) = "ﬁÌ„… «·Ê«—œ"
    
    .TextMatrix(0, 7) = "ﬂ„Ì… ’«œ—"
    .TextMatrix(0, 8) = "ﬁÌ„… ’«œ—"
    
    .TextMatrix(0, 9) = "’«›Ï «·ﬂ„Ì…"
    .TextMatrix(0, 10) = "’«›Ï «·ﬁÌ„…"
    
    .TextMatrix(0, 11) = "‰”»… „»Ì⁄« "
    
    .TextMatrix(0, 12) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 13) = " ﬂ·›… „»Ì⁄« "
    
    .TextMatrix(0, 14) = "—’Ìœ ⁄œœ"
    .TextMatrix(0, 15) = "«·ﬁÌ„…(›—ﬁ Ê«—œ-’«œ—)"
        
    .TextMatrix(0, 16) = "«· ﬂ·›…"
        
    .ColHidden(6) = Not bOpt10
    .ColHidden(8) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    .ColHidden(13) = Not bOpt10
    .ColHidden(15) = Not bOpt10
    
    .ColWidth(0) = 1300
    .ColWidth(1) = 0
    .ColWidth(2) = 600
    .ColWidth(3) = 1500
    .ColWidth(4) = 2000
    
    .ColWidth(5) = 800
    .ColWidth(6) = 1100
    .ColWidth(7) = 800
    .ColWidth(8) = 1100
    .ColWidth(9) = 800
    .ColWidth(10) = 1100
    .ColWidth(11) = 800
    
    .ColWidth(12) = 800
    .ColWidth(13) = 1100
    
    .ColWidth(14) = 800
    .ColWidth(15) = 1100
    .ColWidth(16) = 1100
    .ColHidden(17) = True
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    For nCol = 5 To 9 Step 2
        .ColFormat(nCol) = "#0"
        .ColFormat(nCol + 1) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        .ColDataType(nCol + 1) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol + 1) = Format(.TextMatrix(nRow, nCol + 1), "#0.00")
        Next nRow
    Next nCol
    .ColFormat(12) = "#0"
    .ColFormat(13) = "#0.00"
    .ColDataType(12) = flexDTDouble
    .ColDataType(13) = flexDTDouble
    
    .ColFormat(14) = "#0"
    .ColFormat(15) = "#0.00"
    .ColDataType(14) = flexDTDouble
    .ColDataType(15) = flexDTDouble
    
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 9)) <> 0 Then .TextMatrix(nRow, 11) = Round((Val(.TextMatrix(nRow, 12)) / Val(.TextMatrix(nRow, 9))) * 100, 2)
        If Val(.TextMatrix(nRow, 15)) <> Val(.TextMatrix(nRow, 16)) Then .Cell(flexcpBackColor, nRow, 15, nRow, 16) = vbGreen
    Next nRow
    .SubtotalPosition = flexSTAbove
    For I = 5 To .Cols - 1
        .Subtotal flexSTSum, -1, I, "#0", vbRed, vbYellow, True, "  "
    Next I
    .Subtotal flexSTCount, -1, 3, "#0", vbRed, vbYellow, True, "  "
    If .Rows > 2 Then
        nRow = 2
        If Val(.TextMatrix(nRow, 9)) <> 0 Then .TextMatrix(nRow, 11) = Round((Val(.TextMatrix(nRow, 12)) / Val(.TextMatrix(nRow, 9))) * 100, 2)
    End If
    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    closeCon con
End Sub

Private Sub GRID7_DBLClick()
    VsAllTrans.cGrModel = grid7.TextMatrix(grid7.Row, 3)
    VsAllTrans.Show
End Sub

Private Sub xModelFact_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    ModelLookupAll Me, oSearchItem
End If
End Sub
Private Sub xGRMODEL_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    GrModelLookupAll Me, oSearchItem
End If
End Sub
Private Sub xSection_LostFocus()
If XSECTION.BoundText = "" Then
    Set DATA3.Recordset = myCmd("Select Code,DescA From File1_50 ORDER BY DESCA", con)
Else
    Set DATA3.Recordset = myCmd("Select Code,DescA From File1_50 where [group] = " & Val(XSECTION.BoundText) & " ORDER BY DESCA", con)
End If
Set xGroup.RowSource = DATA3
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"
End Sub
Private Sub xSupp_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then SuppLookupAll Me, oSearch
End Sub
Sub myProc()
If ActiveControl.Name = xModelFact.Name Then
    xModelFact.text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 2)
    Unload oSearchItem
End If
If ActiveControl.Name = XGRMODEL.Name Then
    XGRMODEL.text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    Unload oSearchItem
End If

If ActiveControl.Name = xSupp.Name Then
    xSupp.BoundText = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    oSearch.Hide
End If
If ActiveControl.Name = xFact.Name Then
    xFact.BoundText = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
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
Private Sub MYLOAD1()
Dim cString  As String
Dim cWhere As String, cWhereField As String, cField As String
With grid1
    
    If IsDate(xDate1.Value) Then
        cWhereField = " AND DATE >= " & DateSq(xDate1.Value)
    End If
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(FILE1_11_ALL.[IN] * FILE1_11_ALL.COST ) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST ) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(([IN]-[OUT]) * FILE1_11_ALL.COST ) ")
    
    cField = cField & ",0," & _
            myiif("(type = '6'   )" & cWhereField, "([OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '6'   )" & cWhereField, "([OUT] * file1_11_all.cost ) ")
        
    If SEC(3).Value Then
        cString = " SELECT  FILE4_10.CODE , FILE4_10.DESCA , FILE1_10.FACT, FACT.DESCA AS factdesca , ' ' AS N4 " & _
                  cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL) " & _
                  " FROM  FILE1_10 INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE" & _
                  " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
                  " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                  " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                  " WHERE file1_10.ISNOITEM = 0 "
    ElseIf SEC(0).Value Then
        cString = " SELECT  FILE4_10.CODE , FILE4_10.DESCA , FILE1_10.[MOSM], FILE1_10.[MOSM] , ' ' AS N4 " & _
                    cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL)" & _
                " FROM  FILE1_10 INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE" & _
                " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
                " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                " WHERE file1_10.ISNOITEM = 0 "
    ElseIf SEC(1).Value Then
        cString = " SELECT  FILE4_10.CODE , FILE4_10.DESCA , FILE1_10.MODELSEX , FILE0_81.DESCA AS factdesca , ' ' AS N4 " & _
                    cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL)" & _
                  " FROM  FILE1_10 INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE" & _
                  " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
                  " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                  " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX" & _
                  " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                  " WHERE file1_10.ISNOITEM = 0 "
    ElseIf SEC(2).Value Then
        cString = " SELECT  FILE4_10.CODE , FILE4_10.DESCA , FILE1_10.MODELAGE, FILE0_82.DESCA , ' ' AS N4 " & _
                    cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL)  " & _
                " FROM  FILE1_10 INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE" & _
                " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
                " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                " LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE" & _
                " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                " WHERE file1_10.ISNOITEM = 0 "
    End If
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    If xMosm.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    If xmosm20.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm20.text)
    
    If XSECTION.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xStore.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " [store] = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If XOutlet.Value = 1 Then cWhere = cWhere & Tr(cWhere) & " Outlet = 1 "
    
    If xage.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)

    If IsDate(xDate2.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xDate2.Value)
    
    If xall(0).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 1 "
    ElseIf xall(1).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 0 "
    End If
            
    If cWhere <> "" Then
        cString = cString & " AND " & cWhere
    End If
        
    If SEC(0).Value Then
        cString = cString & " GROUP BY FILE4_10.CODE , FILE4_10.DESCA , FILE1_10.MOSM ORDER BY FILE4_10.CODE "
    ElseIf SEC(1).Value Then
        cString = cString & " GROUP BY FILE4_10.CODE , FILE4_10.DESCA , FILE1_10.MODELSEX , FILE0_81.DESCA ORDER BY FILE4_10.CODE "
    ElseIf SEC(2).Value Then
        cString = cString & " GROUP BY FILE4_10.CODE , FILE4_10.DESCA , FILE1_10.MODELAGE, FILE0_82.DESCA  ORDER BY FILE4_10.CODE "
    ElseIf SEC(3).Value Then
        cString = cString & " GROUP BY FILE4_10.CODE , FILE4_10.DESCA , FILE1_10.FACT, FACT.DESCA ORDER BY FILE4_10.CODE "
    End If
    Set DATA11.Recordset = myCmd(cString, con)
End With
FixGrid1
If grid1.Rows > 0 Then grid1.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FixGrid1()
With grid1
    .Cols = 17
    .RowHeight(0) = 1000
    .RowHidden(1) = True
    .WordWrap = True
    
    .ColHidden(0) = True
    .ColHidden(2) = True
    .ColHidden(4) = True
    
    .TextMatrix(0, 1) = "«·„Ê—œ"
    
        
    If SEC(3).Value Then .TextMatrix(0, 3) = "«·„’‰⁄"
    If SEC(2).Value Then .TextMatrix(0, 3) = "›∆… ⁄„—Ì…"
    If SEC(1).Value Then .TextMatrix(0, 3) = "«·‰Ê⁄"
    If SEC(0).Value Then .TextMatrix(0, 3) = "«·„Ê”„"
    
    .TextMatrix(0, 5) = "ﬂ„Ì… Ê«—œ"
    .TextMatrix(0, 6) = "ﬁÌ„… «·Ê«—œ"
    
    .TextMatrix(0, 7) = "ﬂ„Ì… ’«œ—"
    .TextMatrix(0, 8) = "ﬁÌ„… ’«œ—"
    
    .TextMatrix(0, 9) = "’«›Ï «·ﬂ„Ì…"
    .TextMatrix(0, 10) = "’«›Ï «·ﬁÌ„…"
    
    .TextMatrix(0, 11) = "‰”»… „»Ì⁄« "
    
    .TextMatrix(0, 12) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 13) = " ﬂ·›… „»Ì⁄« "
    
    .TextMatrix(0, 14) = "—’Ìœ ⁄œœ"
    .TextMatrix(0, 15) = "ﬁÌ„… «·—’Ìœ"
    .TextMatrix(0, 16) = "⁄œœ „ÊœÌ·« "
    
    .ColHidden(6) = Not bOpt10
    .ColHidden(8) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    .ColHidden(13) = Not bOpt10
    .ColHidden(15) = Not bOpt10
    
    .ColWidth(0) = 1300
    .ColWidth(1) = 1500
    .ColWidth(2) = 900
    .ColWidth(3) = 1800
    .ColWidth(4) = 1800
    
    .ColWidth(5) = 900
    .ColWidth(6) = 1200
    .ColWidth(7) = 900
    .ColWidth(8) = 1200
    .ColWidth(9) = 900
    .ColWidth(10) = 1200
    .ColWidth(11) = 900
    
    .ColWidth(12) = 900
    .ColWidth(13) = 1200
    
    .ColWidth(14) = 900
    .ColWidth(15) = 1200
    .ColWidth(16) = 900
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    For nCol = 5 To 9 Step 2
        .ColFormat(nCol) = "#0"
        .ColFormat(nCol + 1) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        .ColDataType(nCol + 1) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol + 1) = Format(.TextMatrix(nRow, nCol + 1), "#0.00")
        Next nRow
    Next nCol
    .ColFormat(12) = "#0"
    .ColFormat(13) = "#0.00"
    .ColDataType(12) = flexDTDouble
    .ColDataType(13) = flexDTDouble
    
    .ColFormat(14) = "#0"
    .ColFormat(15) = "#0.00"
    .ColDataType(14) = flexDTDouble
    .ColDataType(15) = flexDTDouble
    
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 9)) <> 0 Then .TextMatrix(nRow, 11) = Round((Val(.TextMatrix(nRow, 12)) / Val(.TextMatrix(nRow, 9))) * 100, 2)
    Next nRow
    .SubtotalPosition = flexSTAbove
    For I = 5 To .Cols - 1
        .Subtotal flexSTSum, -1, I, "#0", vbRed, vbYellow, True, "  "
    Next I
    nRow = 2
    If .Rows > 2 Then
        If Val(.TextMatrix(nRow, 9)) <> 0 Then .TextMatrix(nRow, 11) = Round((Val(.TextMatrix(nRow, 12)) / Val(.TextMatrix(nRow, 9))) * 100, 2)
    End If
    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub
Private Sub MYLOAD2()
Dim I As Double
Dim cString  As String, cField As String
Dim cWhere As String, cWhereField As String
With GRID2
    If IsDate(xDate1.Value) Then
        cWhereField = " AND DATE >= " & DateSq(xDate1.Value)
    End If
'                           0               1                 2                3               4                5
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(FILE1_11_ALL.[IN] * FILE1_11_ALL.COST ) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST ) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(([IN]-[OUT]) * FILE1_11_ALL.COST ) ")
    
    cField = cField & ",0," & _
            myiif("(type = '6'   )" & cWhereField, "([OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '6'   )" & cWhereField, "([OUT] * file1_11_all.cost ) ")
    
    If SEC(3).Value Then
        cString = " SELECT  FILE1_10SC.CODE , FILE1_10SC.DESCA , ' ' AS N2 , ' ' AS N3 , ' ' AS N4 " & _
                   cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL) " & _
                   " FROM  FILE1_10 INNER JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE" & _
                   " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                   " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                   " WHERE file1_10.ISNOITEM = 0 "
    ElseIf SEC(0).Value Then
        cString = " SELECT  FILE1_10SC.CODE, FILE1_10SC.DESCA , FILE1_10.[MOSM], FILE1_10.[MOSM] , ' ' AS N4 " & _
                    cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL) " & _
                    " FROM  FILE1_10 INNER JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE" & _
                    " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                    " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                    " WHERE file1_10.ISNOITEM = 0 "
    ElseIf SEC(1).Value Then
        cString = " SELECT  FILE1_10SC.CODE , FILE1_10SC.DESCA , FILE1_10.MODELSEX , FILE0_81.DESCA AS factdesca , ' ' AS N4 " & _
                    cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL) " & _
                    " FROM  FILE1_10 INNER JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE" & _
                    " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                    " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX" & _
                    " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                    " WHERE file1_10.ISNOITEM = 0 "
    ElseIf SEC(2).Value Then
        cString = " SELECT  FILE1_10SC.CODE , FILE1_10SC.DESCA , FILE1_10.MODELAGE, FILE0_82.DESCA , ' ' AS N4 " & _
                    cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL) " & _
                    " FROM  FILE1_10 INNER JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE" & _
                    " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                    " LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE" & _
                    " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                    " WHERE file1_10.ISNOITEM = 0 "
    End If
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    If xMosm.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    If xmosm20.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm20.text)
    
    If XSECTION.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xStore.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " [store] = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If XOutlet.Value = 1 Then cWhere = cWhere & Tr(cWhere) & " Outlet = 1 "
    
    If xage.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    

    If IsDate(xDate2.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xDate2.Value)
    
    If xall(0).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 1 "
    ElseIf xall(1).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 0 "
    End If
            
    If cWhere <> "" Then
        cString = cString & " AND " & cWhere
    End If

    If SEC(0).Value Then
        cString = cString & " GROUP BY FILE1_10SC.CODE , FILE1_10SC.DESCA , FILE1_10.MOSM ORDER BY FILE1_10SC.CODE , FILE1_10.MOSM "
    ElseIf SEC(1).Value Then
        cString = cString & " GROUP BY FILE1_10SC.CODE , FILE1_10SC.DESCA , FILE1_10.MODELSEX , FILE0_81.DESCA ORDER BY FILE1_10SC.CODE "
    ElseIf SEC(2).Value Then
        cString = cString & " GROUP BY FILE1_10SC.CODE , FILE1_10SC.DESCA , FILE1_10.MODELAGE, FILE0_82.DESCA  ORDER BY FILE1_10SC.CODE "
    ElseIf SEC(3).Value Then
        cString = cString & " GROUP BY FILE1_10SC.CODE , FILE1_10SC.DESCA ORDER BY FILE1_10SC.CODE "
    End If
    Set data12.Recordset = myCmd(cString, con)
End With
FixGrid2
If GRID2.Rows > 0 Then GRID2.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FixGrid2()
With GRID2
    .Cols = 17
    .RowHeight(0) = 1000
    .RowHidden(1) = True
    .WordWrap = True
    
    .ColHidden(0) = True
    .ColHidden(2) = True
    .ColHidden(4) = True
    
    .TextMatrix(0, 1) = "«·ﬁ”„"
    
        
    If SEC(3).Value Then .TextMatrix(0, 3) = ""
    If SEC(2).Value Then .TextMatrix(0, 3) = "›∆… ⁄„—Ì…"
    If SEC(1).Value Then .TextMatrix(0, 3) = "«·‰Ê⁄"
    If SEC(0).Value Then .TextMatrix(0, 3) = "«·„Ê”„"
    
    .TextMatrix(0, 5) = "ﬂ„Ì… Ê«—œ"
    .TextMatrix(0, 6) = "ﬁÌ„… «·Ê«—œ"
    
    .TextMatrix(0, 7) = "ﬂ„Ì… ’«œ—"
    .TextMatrix(0, 8) = "ﬁÌ„… ’«œ—"
    
    .TextMatrix(0, 9) = "’«›Ï «·ﬂ„Ì…"
    .TextMatrix(0, 10) = "’«›Ï «·ﬁÌ„…"
    
    .TextMatrix(0, 11) = "‰”»… „»Ì⁄« "
    
    .TextMatrix(0, 12) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 13) = " ﬂ·›… „»Ì⁄« "
    
    .TextMatrix(0, 14) = "—’Ìœ ⁄œœ"
    .TextMatrix(0, 15) = "ﬁÌ„… «·—’Ìœ"
    .TextMatrix(0, 16) = "⁄œœ „ÊœÌ·« "
    
    .ColHidden(6) = Not bOpt10
    .ColHidden(8) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    .ColHidden(13) = Not bOpt10
    .ColHidden(15) = Not bOpt10
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1500
    .ColWidth(2) = 900
    .ColWidth(3) = 1500
    .ColWidth(4) = 2100
    
    .ColWidth(5) = 900
    .ColWidth(6) = 1200
    .ColWidth(7) = 900
    .ColWidth(8) = 1200
    .ColWidth(9) = 900
    .ColWidth(10) = 1200
    .ColWidth(11) = 900
    
    .ColWidth(12) = 900
    .ColWidth(13) = 1200
    
    .ColWidth(14) = 900
    .ColWidth(15) = 1200
    .ColWidth(16) = 800
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    For nCol = 5 To 9 Step 2
        .ColFormat(nCol) = "#0"
        .ColFormat(nCol + 1) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        .ColDataType(nCol + 1) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol + 1) = Format(.TextMatrix(nRow, nCol + 1), "#0.00")
        Next nRow
    Next nCol
    .ColFormat(12) = "#0"
    .ColFormat(13) = "#0.00"
    .ColDataType(12) = flexDTDouble
    .ColDataType(13) = flexDTDouble
    
    .ColFormat(14) = "#0"
    .ColFormat(15) = "#0.00"
    .ColDataType(14) = flexDTDouble
    .ColDataType(15) = flexDTDouble
    
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 9)) <> 0 Then .TextMatrix(nRow, 11) = Round((Val(.TextMatrix(nRow, 12)) / Val(.TextMatrix(nRow, 9))) * 100, 2)
    Next nRow
    .SubtotalPosition = flexSTAbove
    For I = 5 To .Cols - 1
        .Subtotal flexSTSum, -1, I, "#0", vbRed, vbYellow, True, "  "
    Next I
    If .Rows > 2 Then
        nRow = 2
        If Val(.TextMatrix(nRow, 9)) <> 0 Then .TextMatrix(nRow, 11) = Round((Val(.TextMatrix(nRow, 12)) / Val(.TextMatrix(nRow, 9))) * 100, 2)
    End If
    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub
Private Sub MYLOAD3()
Dim I As Double, cCountM As String
Dim cString  As String
Dim cField As String, cWhere As String, cWhereField As String
With GRID3

    If IsDate(xDate1.Value) Then
        cWhereField = " AND DATE >= " & DateSq(xDate1.Value)
    End If
        
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f')" & cWhereField, "([IN]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f')" & cWhereField, "(FILE1_11_ALL.[IN] * FILE1_11_ALL.COST ) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f')" & cWhereField, "([OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f')" & cWhereField, "(FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST ) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f')" & cWhereField, "([IN]-[OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f')" & cWhereField, "(([IN]-[OUT]) * FILE1_11_ALL.COST ) ")
    
    cField = cField & ",0," & _
            myiif("(type = '6')" & cWhereField, "([OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '6')" & cWhereField, "([OUT] * file1_11_all.cost ) ")
        
    cWhereField = ""
    
    If xMosm.MatchedWithList Then
        cWhereField = cWhereField & " AND FILE1_10_2.MOSM = " & MyParn(xMosm.BoundText)
    End If
    
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then
            cWhereField = cWhereField & " AND ( SUBSTRING(FILE1_10_2.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10_2.MOSM  = '000'  )"
        End If
    Else
        If XMOSM2.text <> "" Then
            cWhereField = cWhereField & " AND ( SUBSTRING(FILE1_10_2.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10_2.MOSM ,1,1) = 'M'  OR FILE1_10_2.MOSM  = '000' ) "
        End If
    End If
    
    If xmosm20.text <> "" Then
        cWhereField = cWhereField & " AND FILE1_10_2.MOSM2 = " & MyParn(xmosm20.text)
    End If
    
    If XSECTION.MatchedWithList Then cWhereField = cWhereField & " AND FILE1_10_2.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhereField = cWhereField & " AND FILE1_10_2.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhereField = cWhereField & " AND FILE1_10_2.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhereField = cWhereField & " AND FILE1_10_2.FACT = " & MyParn(xFact.BoundText)
    If XOutlet.Value = 1 Then cWhereField = cWhereField & " AND Outlet = 1 "
    If xage.MatchedWithList Then cWhereField = cWhereField & " AND FILE1_10_2.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhereField = cWhereField & " AND FILE1_10_2.MODELSEX = " & MyParn(xsex.BoundText)
    
    cCountM = " ,(select COUNT(DISTINCT FILE1_10_2.MODEL) from file1_10 as file1_10_2 where ISNODEM = 0   and file1_10_2.[group] = file1_50.code " & cWhereField & " )"
    
    If SEC(3).Value Then
        cString = " SELECT  FILE1_50.CODE , FILE1_50.DESCA , ' ' AS N2 , ' ' AS N3 , ' ' AS N4 " & _
                    cField & ", SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL) " & cCountM & _
                    " FROM  FILE1_10 INNER JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.[CODE]" & _
                    " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                    " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                    " WHERE file1_10.ISNOITEM = 0 "
    ElseIf SEC(0).Value Then
        cString = " SELECT  FILE1_50.CODE , FILE1_50.DESCA , FILE1_10.[MOSM], FILE1_10.[MOSM] , ' ' AS N4 " & _
                cField & " ,SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL) " & cCountM & _
                " FROM  FILE1_10 INNER JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.[CODE]" & _
                " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                " WHERE file1_10.ISNOITEM = 0 "
    ElseIf SEC(1).Value Then
        cString = " SELECT  FILE1_50.CODE , FILE1_50.DESCA , FILE1_10.MODELSEX , FILE0_81.DESCA AS factdesca , ' ' AS N4 " & _
                cField & ",SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL) " & cCountM & _
                " FROM  FILE1_10 INNER JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.[CODE]" & _
                " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX" & _
                " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                " WHERE file1_10.ISNOITEM = 0 "
    ElseIf SEC(2).Value Then
        cString = " SELECT  FILE1_50.CODE , FILE1_50.DESCA , FILE1_10.MODELAGE, FILE0_82.DESCA , ' ' AS N4 " & _
                    cField & ", SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL)  " & cCountM & _
                    " FROM  FILE1_10 INNER JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.[CODE]" & _
                    " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                    " LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE" & _
                    " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                    " WHERE file1_10.ISNOITEM = 0 "
    End If
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    If xMosm.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    If xmosm20.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm20.text)
    
    
    If XSECTION.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xStore.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " [store] = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If XOutlet.Value = 1 Then cWhere = cWhere & Tr(cWhere) & " Outlet = 1 "
    
    If xage.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)

    If IsDate(xDate2.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xDate2.Value)
    
    If xall(0).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 1 "
    ElseIf xall(1).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 0 "
    End If
    
    If cWhere <> "" Then cString = cString & " AND " & cWhere
    
    If SEC(0).Value Then
        cString = cString & " GROUP BY FILE1_50.CODE , FILE1_50.DESCA , FILE1_10.MOSM ORDER BY FILE1_50.CODE "
    ElseIf SEC(1).Value Then
        cString = cString & " GROUP BY FILE1_50.CODE , FILE1_50.DESCA , FILE1_10.MODELSEX , FILE0_81.DESCA ORDER BY FILE1_50.CODE "
    ElseIf SEC(2).Value Then
        cString = cString & " GROUP BY FILE1_50.CODE , FILE1_50.DESCA , FILE1_10.MODELAGE, FILE0_82.DESCA  ORDER BY FILE1_50.CODE "
    ElseIf SEC(3).Value Then
        cString = cString & " GROUP BY FILE1_50.CODE , FILE1_50.DESCA ORDER BY FILE1_50.CODE "
    End If
    Set DATA13.Recordset = myCmd(cString, con)
End With
FixGrid3
If GRID3.Rows > 0 Then GRID3.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FixGrid3()
With GRID3
    .Cols = 18
    .RowHeight(0) = 1000
    .RowHidden(1) = True
    .WordWrap = True
    
    .ColHidden(0) = True
    .ColHidden(2) = True
    .ColHidden(4) = True
    
    .TextMatrix(0, 1) = "«·„Ã„Ê⁄…"
        
    If SEC(3).Value Then
        .TextMatrix(0, 3) = ""
        .ColHidden(3) = True
    End If
    
    If SEC(2).Value Then .TextMatrix(0, 3) = "›∆… ⁄„—Ì…"
    If SEC(1).Value Then .TextMatrix(0, 3) = "«·‰Ê⁄"
    If SEC(0).Value Then .TextMatrix(0, 3) = "«·„Ê”„"
    
    .TextMatrix(0, 5) = "ﬂ„Ì… Ê«—œ"
    .TextMatrix(0, 6) = "ﬁÌ„… «·Ê«—œ"
    
    .TextMatrix(0, 7) = "ﬂ„Ì… ’«œ—"
    .TextMatrix(0, 8) = "ﬁÌ„… ’«œ—"
    
    .TextMatrix(0, 9) = "’«›Ï «·ﬂ„Ì…"
    .TextMatrix(0, 10) = "’«›Ï «·ﬁÌ„…"
    
    .TextMatrix(0, 11) = "‰”»… „»Ì⁄« "
    
    .TextMatrix(0, 12) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 13) = " ﬂ·›… „»Ì⁄« "
    
    .TextMatrix(0, 14) = "—’Ìœ ⁄œœ"
    .TextMatrix(0, 15) = "ﬁÌ„… «·—’Ìœ"
    .TextMatrix(0, 16) = "⁄œœ „ÊœÌ·« "
    .TextMatrix(0, 17) = "„ÊœÌ·«  „”„ÊÕ ·Â« ÿ·»Ì« "
    
    .ColHidden(6) = Not bOpt10
    .ColHidden(8) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    .ColHidden(13) = Not bOpt10
    .ColHidden(15) = Not bOpt10
    
    .ColWidth(0) = 1300
    .ColWidth(1) = 1500
    .ColWidth(2) = 900
    .ColWidth(3) = 1500
    .ColWidth(4) = 2100
    
    .ColWidth(5) = 900
    .ColWidth(6) = 1200
    .ColWidth(7) = 900
    .ColWidth(8) = 1200
    .ColWidth(9) = 900
    .ColWidth(10) = 1200
    .ColWidth(11) = 900
    
    .ColWidth(12) = 900
    .ColWidth(13) = 1200
    
    .ColWidth(14) = 900
    .ColWidth(15) = 1200
    .ColWidth(16) = 800
    .ColWidth(17) = 1000
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    For nCol = 5 To 9 Step 2
        .ColFormat(nCol) = "#0"
        .ColFormat(nCol + 1) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        .ColDataType(nCol + 1) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol + 1) = Format(.TextMatrix(nRow, nCol + 1), "#0.00")
        Next nRow
    Next nCol
    .ColFormat(12) = "#0"
    .ColFormat(13) = "#0.00"
    .ColDataType(12) = flexDTDouble
    .ColDataType(13) = flexDTDouble
    
    .ColFormat(14) = "#0"
    .ColFormat(15) = "#0.00"
    .ColDataType(14) = flexDTDouble
    .ColDataType(15) = flexDTDouble
    
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 9)) <> 0 Then .TextMatrix(nRow, 11) = Round((Val(.TextMatrix(nRow, 12)) / Val(.TextMatrix(nRow, 9))) * 100, 2)
    Next nRow
    .SubtotalPosition = flexSTAbove
    For I = 5 To .Cols - 1
        .Subtotal flexSTSum, -1, I, "#0", vbRed, vbYellow, True, "  "
    Next I
    If .Rows > 2 Then
        nRow = 2
        If Val(.TextMatrix(nRow, 9)) <> 0 Then .TextMatrix(nRow, 11) = Round((Val(.TextMatrix(nRow, 12)) / Val(.TextMatrix(nRow, 9))) * 100, 2)
    End If
    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub
Private Sub MYLOAD4()
Dim cString  As String
Dim cField As String, cFieldCount As String, cWhere As String, cWhereField As String
With GRID4
    If IsDate(xDate1.Value) Then
        cWhereField = " AND DATE >= " & DateSq(xDate1.Value)
    End If

    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(FILE1_11_ALL.[IN] * FILE1_11_ALL.COST ) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST ) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(([IN]-[OUT]) * FILE1_11_ALL.COST ) ")
    
    cField = cField & ",0," & _
            myiif("(type = '6')" & cWhereField, "([OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '6'   )" & cWhereField, "([OUT] * file1_11_all.cost ) ")

    ' Count Field
        
    cWhereField = ""
    If xMosm.MatchedWithList Then cWhereField = cWhereField & " AND FILE1_10_2.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhereField = cWhereField & " AND ( SUBSTRING(FILE1_10_2.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10_2.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhereField = cWhereField & " AND ( SUBSTRING(FILE1_10_2.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10_2.MOSM ,1,1) = 'M'  OR FILE1_10_2.MOSM  = '000' ) "
    End If
    If xmosm20.text <> "" Then cWhereField = cWhereField & " AND FILE1_10_2.MOSM2 = " & MyParn(xmosm20.text)
    If XSECTION.MatchedWithList Then cWhereField = cWhereField & " AND FILE1_10_2.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhereField = cWhereField & " AND FILE1_10_2.[CODE] = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhereField = cWhereField & " AND FILE1_10_2.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhereField = cWhereField & " AND FILE1_10_2.FACT = " & MyParn(xFact.BoundText)
    If XOutlet.Value = 1 Then cWhereField = cWhereField & " AND Outlet = 1 "
    If xage.MatchedWithList Then cWhereField = cWhereField & " AND FILE1_10_2.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhereField = cWhereField & " AND FILE1_10_2.MODELSEX = " & MyParn(xsex.BoundText)
    
        
    If SEC(3).Value Then
        cFieldCount = ",(SELECT COUNT(DISTINCT FILE1_10_2.MODEL) FROM FILE1_10 AS FILE1_10_2 INNER JOIN FILE1_11_ALL AS FILE1_11_ALL_2 ON FILE1_10_2.ITEM = FILE1_11_ALL_2.ITEM WHERE FILE1_10_2.ISNODEM = 0 AND FILE1_10_2.MODELAGE = FILE0_82.CODE " & cWhereField & " )"
        
        cString = " SELECT  FILE0_82.CODE , FILE0_82.DESCA , ' ' AS N2 , ' ' AS N3 , ' ' AS N4 " & _
                    cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL)" & cFieldCount & _
                    " FROM  FILE1_10 INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                    " LEFT OUTER JOIN FILE0_82 ON FILE1_10.MODELAGE = FILE0_82.CODE" & _
                    " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                    " WHERE file1_10.ISNOITEM = 0 "
    ElseIf SEC(0).Value Then
        cFieldCount = ",(SELECT COUNT(DISTINCT FILE1_10_2.MODEL) FROM FILE1_10 AS FILE1_10_2 INNER JOIN FILE1_11_ALL AS FILE1_11_ALL_2 ON FILE1_10_2.ITEM = FILE1_11_ALL_2.ITEM WHERE FILE1_10_2.ISNODEM = 0 AND FILE1_10_2.MODELAGE = FILE0_82.CODE AND FILE1_10.MOSM = FILE1_10_2.MOSM " & cWhereField & " )    "
        
        cString = " SELECT  FILE0_82.CODE , FILE0_82.DESCA , FILE1_10.[MOSM], FILE1_10.[MOSM] , ' ' AS N4 " & _
                    cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL)" & cFieldCount & _
                    " FROM  FILE1_10 LEFT OUTER JOIN FILE0_82 ON FILE1_10.MODELAGE = FILE0_82.CODE" & _
                    " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                    " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                    " WHERE file1_10.ISNOITEM = 0 "
    
    ElseIf SEC(1).Value Then
        cFieldCount = ",(SELECT COUNT(DISTINCT FILE1_10_2.MODEL) FROM FILE1_10 AS FILE1_10_2 INNER JOIN FILE1_11_ALL AS FILE1_11_ALL_2 ON FILE1_10_2.ITEM = FILE1_11_ALL_2.ITEM WHERE FILE1_10_2.ISNODEM = 0 AND FILE1_10_2.MODELAGE = FILE0_82.CODE AND FILE1_10.MODELSEX = FILE1_10_2.MODELSEX " & cWhereField & ")"
        cString = " SELECT  FILE0_82.CODE , FILE0_82.DESCA , FILE1_10.MODELSEX , FILE0_81.DESCA AS factdesca , ' ' AS N4 " & _
                    cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL)" & cFieldCount & _
                    " FROM  file1_10 LEFT OUTER JOIN FILE0_82 ON FILE1_10.MODELAGE = FILE0_82.CODE" & _
                    " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                    " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX" & _
                    " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                    " WHERE file1_10.ISNOITEM = 0 "
    ElseIf SEC(2).Value Then
        cFieldCount = ",(SELECT COUNT(DISTINCT FILE1_10_2.MODEL) FROM FILE1_10 AS FILE1_10_2 INNER JOIN FILE1_11_ALL AS FILE1_11_ALL_2 ON FILE1_10_2.ITEM = FILE1_11_ALL_2.ITEM WHERE FILE1_10_2.ISNODEM = 0 AND FILE1_10_2.MODELAGE = FILE0_82.CODE " & cWhere2 & " )    "
        cString = " SELECT  FILE0_82.CODE , FILE0_82.DESCA , FILE0_82.CODE , FILE0_82.DESCA , ' ' AS N4 " & _
                    cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL)" & cFieldCount & _
                    " FROM  FILE1_10 LEFT OUTER JOIN FILE0_82 ON FILE1_10.MODELAGE = FILE0_82.CODE" & _
                    " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                    " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                    " WHERE file1_10.ISNOITEM = 0 "
    End If
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    If xMosm.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    If xmosm20.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm20.text)
    If XSECTION.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If XOutlet.Value = 1 Then cWhere = cWhere & Tr(cWhere) & " Outlet = 1 "
    If xage.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xSupp.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[CODE] = " & MyParn(xSupp.BoundText)
    If xStore.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " [store] = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "

    If IsDate(xDate2.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xDate2.Value)
    
    If xall(0).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 1 "
    ElseIf xall(1).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 0 "
    End If
    
    If cWhere <> "" Then cString = cString & " AND " & cWhere
    If SEC(0).Value Then
        cString = cString & " GROUP BY FILE0_82.CODE , FILE0_82.DESCA , FILE1_10.MOSM ORDER BY FILE0_82.CODE , FILE1_10.MOSM "
    ElseIf SEC(1).Value Then
        cString = cString & " GROUP BY FILE0_82.CODE , FILE0_82.DESCA , FILE1_10.MODELSEX , FILE0_81.DESCA ORDER BY FILE0_82.CODE , FILE1_10.MODELSEX "
    ElseIf SEC(2).Value Then
        cString = cString & " GROUP BY FILE0_82.CODE , FILE0_82.DESCA ORDER BY FILE0_82.CODE "
    ElseIf SEC(3).Value Then
        cString = cString & " GROUP BY FILE0_82.CODE , FILE0_82.DESCA ORDER BY FILE0_82.CODE "
    End If
    Set DATA14.Recordset = myCmd(cString, con)
End With
FIXGRID4
If GRID4.Rows > 0 Then GRID4.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FIXGRID4()
With GRID4
    .Cols = 18
    .RowHeight(0) = 1000
    .RowHidden(1) = True
    .WordWrap = True
    
    .ColHidden(0) = True
    .ColHidden(2) = True
    .ColHidden(4) = True
    
    .TextMatrix(0, 1) = "«·›∆… «·⁄„—Ì…"
    
        
    If SEC(3).Value Then
        .TextMatrix(0, 3) = ""
        .ColHidden(3) = True
    End If
    If SEC(2).Value Then .TextMatrix(0, 3) = "›∆… ⁄„—Ì…"
    If SEC(1).Value Then .TextMatrix(0, 3) = "«·‰Ê⁄"
    If SEC(0).Value Then .TextMatrix(0, 3) = "«·„Ê”„"
    
    .TextMatrix(0, 5) = "ﬂ„Ì… Ê«—œ"
    .TextMatrix(0, 6) = "ﬁÌ„… «·Ê«—œ"
    
    .TextMatrix(0, 7) = "ﬂ„Ì… ’«œ—"
    .TextMatrix(0, 8) = "ﬁÌ„… ’«œ—"
    
    .TextMatrix(0, 9) = "’«›Ï «·ﬂ„Ì…"
    .TextMatrix(0, 10) = "’«›Ï «·ﬁÌ„…"
    
    .TextMatrix(0, 11) = "‰”»… „»Ì⁄« "
    
    .TextMatrix(0, 12) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 13) = " ﬂ·›… „»Ì⁄« "
    
    .TextMatrix(0, 14) = "—’Ìœ ⁄œœ"
    .TextMatrix(0, 15) = "ﬁÌ„… «·—’Ìœ"
    .TextMatrix(0, 16) = "⁄œœ „ÊœÌ·« "
    
    .TextMatrix(0, 17) = "⁄œœ „ÊœÌ·«  „ «Õ ·œÏ «·„Ê—œ"
    
    .ColHidden(6) = Not bOpt10
    .ColHidden(8) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    .ColHidden(13) = Not bOpt10
    .ColHidden(15) = Not bOpt10
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1500
    .ColWidth(2) = 900
    .ColWidth(3) = 1500
    .ColWidth(4) = 2100
    
    .ColWidth(5) = 900
    .ColWidth(6) = 1200
    .ColWidth(7) = 900
    .ColWidth(8) = 1200
    .ColWidth(9) = 900
    .ColWidth(10) = 1200
    .ColWidth(11) = 900
    
    .ColWidth(12) = 900
    .ColWidth(13) = 1200
    
    .ColWidth(14) = 900
    .ColWidth(15) = 1200
    .ColWidth(16) = 800
    .ColWidth(16) = 1000
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    For nCol = 5 To 9 Step 2
        .ColFormat(nCol) = "#0"
        .ColFormat(nCol + 1) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        .ColDataType(nCol + 1) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol + 1) = Format(.TextMatrix(nRow, nCol + 1), "#0.00")
        Next nRow
    Next nCol
    .ColFormat(12) = "#0"
    .ColFormat(13) = "#0.00"
    .ColDataType(12) = flexDTDouble
    .ColDataType(13) = flexDTDouble
    
    .ColFormat(14) = "#0"
    .ColFormat(15) = "#0.00"
    .ColDataType(14) = flexDTDouble
    .ColDataType(15) = flexDTDouble
    .ColFormat(17) = "#0"
    
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 9)) <> 0 Then .TextMatrix(nRow, 11) = Round((Val(.TextMatrix(nRow, 12)) / Val(.TextMatrix(nRow, 9))) * 100, 2)
    Next nRow
    .SubtotalPosition = flexSTAbove
    For I = 5 To .Cols - 1
        .Subtotal flexSTSum, -1, I, "#0", vbRed, vbYellow, True, "  "
    Next I
    If .Rows > 2 Then
        nRow = 2
        If Val(.TextMatrix(nRow, 9)) <> 0 Then .TextMatrix(nRow, 11) = Round((Val(.TextMatrix(nRow, 12)) / Val(.TextMatrix(nRow, 9))) * 100, 2)
    End If
    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub
Private Sub myload6()
Dim cString  As String
Dim cField As String, cWhere As String, cWhereField As String
With grid6
    If IsDate(xDate1.Value) Then
        cWhereField = " AND DATE >= " & DateSq(xDate1.Value)
    End If
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(FILE1_11_ALL.[IN] * FILE1_11_ALL.COST ) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST ) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(([IN]-[OUT]) * FILE1_11_ALL.COST ) ")
    
    cField = cField & ",0," & _
            myiif("(type = '6' )" & cWhereField, "([OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '6' )" & cWhereField, "([OUT] * file1_11_all.cost ) ")
    
    
    If SEC(3).Value Then
        cString = " SELECT  STORE_ALL.CODE , STORE_ALL.DESCA , ' ' AS N2 , ' ' AS N3 , ' ' AS N4 " & _
                    cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL)" & _
                    " FROM  FILE1_10 INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                    " INNER JOIN STORE_ALL ON STORE_ALL.CODE = FILE1_11_ALL.STORE" & _
                    " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                    " WHERE file1_10.ISNOITEM = 0 "
    ElseIf SEC(0).Value Then
        cString = " SELECT  STORE_ALL.CODE , STORE_ALL.DESCA , FILE1_10.[MOSM], FILE1_10.[MOSM] , ' ' AS N4 " & _
                    cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL)" & _
                    " FROM  FILE1_10 INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                    " INNER JOIN STORE_ALL ON STORE_ALL.CODE = FILE1_11_ALL.STORE" & _
                    " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                    " WHERE file1_10.ISNOITEM = 0 "
    ElseIf SEC(1).Value Then
        cString = " SELECT  STORE_ALL.CODE , STORE_ALL.DESCA, FILE1_10.MODELSEX , FILE0_81.DESCA AS factdesca , ' ' AS N4 " & _
                cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL)" & _
                " FROM  FILE1_10 INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                " INNER JOIN STORE_ALL ON STORE_ALL.CODE = FILE1_11_ALL.STORE" & _
                " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX" & _
                " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                " WHERE file1_10.ISNOITEM = 0 "
    ElseIf SEC(2).Value Then
        cString = " SELECT  STORE_ALL.CODE , STORE_ALL.DESCA , FILE1_10.MODELAGE, FILE0_82.DESCA , ' ' AS N4 " & _
                    cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL)" & _
                    " FROM  FILE1_10 INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                    " INNER JOIN STORE_ALL ON STORE_ALL.CODE = FILE1_11_ALL.STORE" & _
                    " LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE" & _
                    " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                    " WHERE file1_10.ISNOITEM = 0 "
    End If
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    If xMosm.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    If xmosm20.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm20.text)
    
    If XSECTION.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xStore.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " [store] = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    
    If xage.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If XOutlet.Value = 1 Then cWhere = cWhere & Tr(cWhere) & " Outlet = 1 "

    If IsDate(xDate2.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xDate2.Value)
    
    If xall(0).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 1 "
    ElseIf xall(1).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 0 "
    End If
    
    If cWhere <> "" Then cString = cString & " AND " & cWhere
    
    If SEC(0).Value Then
        cString = cString & " GROUP BY STORE_ALL.CODE , STORE_ALL.DESCA , FILE1_10.MOSM ORDER BY STORE_ALL.CODE "
    ElseIf SEC(1).Value Then
        cString = cString & " GROUP BY STORE_ALL.CODE , STORE_ALL.DESCA , FILE1_10.MODELSEX , FILE0_81.DESCA ORDER BY STORE_ALL.CODE "
    ElseIf SEC(2).Value Then
        cString = cString & " GROUP BY STORE_ALL.CODE , STORE_ALL.DESCA , FILE1_10.MODELAGE, FILE0_82.DESCA  ORDER BY STORE_ALL.CODE "
    ElseIf SEC(3).Value Then
        cString = cString & " GROUP BY STORE_ALL.CODE , STORE_ALL.DESCA ORDER BY STORE_ALL.CODE "
    End If
    Set DATA16.Recordset = myCmd(cString, con)
End With
FIXGRID6
If grid6.Rows > 0 Then grid6.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FIXGRID6()
With grid6
    .Cols = 17
    .RowHeight(0) = 1000
    .RowHidden(1) = True
    .WordWrap = True
    
    .ColHidden(0) = True
    .ColHidden(2) = True
    .ColHidden(4) = True
    
    .TextMatrix(0, 1) = "«·›—⁄"
        
    If SEC(3).Value Then
        .TextMatrix(0, 3) = ""
        .ColHidden(3) = True
    End If
    
    If SEC(2).Value Then .TextMatrix(0, 3) = "›∆… ⁄„—Ì…"
    If SEC(1).Value Then .TextMatrix(0, 3) = "«·‰Ê⁄"
    If SEC(0).Value Then .TextMatrix(0, 3) = "«·„Ê”„"
    
    .TextMatrix(0, 5) = "ﬂ„Ì… Ê«—œ"
    .TextMatrix(0, 6) = "ﬁÌ„… «·Ê«—œ"
    
    .TextMatrix(0, 7) = "ﬂ„Ì… ’«œ—"
    .TextMatrix(0, 8) = "ﬁÌ„… ’«œ—"
    
    .TextMatrix(0, 9) = "’«›Ï «·ﬂ„Ì…"
    .TextMatrix(0, 10) = "’«›Ï «·ﬁÌ„…"
    
    .TextMatrix(0, 11) = "‰”»… „»Ì⁄« "
    
    .TextMatrix(0, 12) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 13) = " ﬂ·›… „»Ì⁄« "
    
    .TextMatrix(0, 14) = "—’Ìœ ⁄œœ"
    .TextMatrix(0, 15) = "ﬁÌ„… «·—’Ìœ"
    .TextMatrix(0, 16) = "⁄œœ „ÊœÌ·« "
    
    .ColHidden(6) = Not bOpt10
    .ColHidden(8) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    .ColHidden(13) = Not bOpt10
    .ColHidden(15) = Not bOpt10
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 2500
    .ColWidth(2) = 900
    .ColWidth(3) = 1500
    .ColWidth(4) = 2100
    
    .ColWidth(5) = 900
    .ColWidth(6) = 1200
    .ColWidth(7) = 900
    .ColWidth(8) = 1200
    .ColWidth(9) = 900
    .ColWidth(10) = 1200
    .ColWidth(11) = 900
    
    .ColWidth(12) = 900
    .ColWidth(13) = 1200
    
    .ColWidth(14) = 900
    .ColWidth(15) = 1200
    .ColWidth(16) = 800
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    For nCol = 5 To 9 Step 2
        .ColFormat(nCol) = "#0"
        .ColFormat(nCol + 1) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        .ColDataType(nCol + 1) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol + 1) = Format(.TextMatrix(nRow, nCol + 1), "#0.00")
        Next nRow
    Next nCol
    .ColFormat(12) = "#0"
    .ColFormat(13) = "#0.00"
    .ColDataType(12) = flexDTDouble
    .ColDataType(13) = flexDTDouble
    
    .ColFormat(14) = "#0"
    .ColFormat(15) = "#0.00"
    .ColDataType(14) = flexDTDouble
    .ColDataType(15) = flexDTDouble
    
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 9)) <> 0 Then .TextMatrix(nRow, 11) = Round((Val(.TextMatrix(nRow, 12)) / Val(.TextMatrix(nRow, 9))) * 100, 2)
    Next nRow
    .SubtotalPosition = flexSTAbove
    For I = 5 To .Cols - 1
        .Subtotal flexSTSum, -1, I, "#0", vbRed, vbYellow, True, "  "
    Next I
    If .Rows > 2 Then
        nRow = 2
        If Val(.TextMatrix(nRow, 9)) <> 0 Then .TextMatrix(nRow, 11) = Round((Val(.TextMatrix(nRow, 12)) / Val(.TextMatrix(nRow, 9))) * 100, 2)
    End If
    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub
Private Sub myload7()
Dim cString  As String
Dim cField As String, cWhere As String, cWhereField As String
With grid7
    If IsDate(xDate1.Value) Then
        cWhereField = " AND DATE >= " & DateSq(xDate1.Value)
    End If
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(FILE1_11_ALL.[IN] * FILE1_11_ALL.COST ) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST ) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "([IN]-[OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhereField, "(([IN]-[OUT]) * FILE1_11_ALL.COST ) ")


    cField = cField & ",0," & _
            myiif("(type = '6'   )" & cWhereField, "([OUT]) ")

    cField = cField & "," & _
            myiif("(type = '6'   )" & cWhereField, "([OUT] * FILE1_11_ALL.COST ) ")

    cField2 = "DATEDIFF(dd , " & myiif2("type = '2' ", "[date]", "MIN") & " , GETDATE()) AS DATEDIFF"
    
    If SEC(3).Value Or SEC(0).Value Then
        cString = " SELECT  FILE4_10.DESCA AS supldesca, FACT.DESCA AS factdesca, FILE1_10.MOSM, FILE1_10.GRMODEL , max(FILE1_10.GRDESCA) " & _
                    cField & " , sum([in]-out) , sum(([in]-out) * FILE1_11_ALL.COST) , COUNT(DISTINCT MODEL ) , " & cField2 & _
                    " FROM  FILE1_10 INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE" & _
                    " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
                    " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                    " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                    " WHERE file1_10.ISNOITEM = 0 "
    ElseIf SEC(2).Value Then
        cString = " SELECT  FILE4_10.DESCA AS supldesca, FACT.DESCA AS factdesca, FILE0_82.DESCA, FILE1_10.GRMODEL , max(FILE1_10.GRDESCA) " & _
                   cField & " , sum([in]-out) , sum(([in]-out) * FILE1_11_ALL.COST) , COUNT(DISTINCT MODEL ) , " & cField2 & _
                   " FROM  FILE1_10 INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM  WHERE file1_10.ISNOITEM = 0 "
    ElseIf SEC(1).Value Then
        cString = " SELECT  FILE4_10.DESCA AS supldesca, FACT.DESCA AS factdesca, FILE0_81.DESCA, FILE1_10.GRMODEL , max(FILE1_10.GRDESCA) " & _
                    cField & " , sum([in]-out) , sum(([in]-out) * FILE1_11_ALL.COST) , COUNT(DISTINCT MODEL ) , " & cField2 & _
                    " FROM  FILE1_10 INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE" & _
                    " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
                    " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX" & _
                    " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                    " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
                    " WHERE file1_10.ISNOITEM = 0 "
    End If
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"

    If xMosm.MatchedWithList Then
        cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    ElseIf XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    If xmosm20.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm20.text)
    If XSECTION.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xStore.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " [store] = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If XOutlet.Value = 1 Then cWhere = cWhere & Tr(cWhere) & " Outlet = 1 "
    
    If xage.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)

    If IsDate(xDate2.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xDate2.Value)
    
    If xall(0).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 1 "
    ElseIf xall(1).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 0 "
    End If
            
    If cWhere <> "" Then cString = cString & " AND " & cWhere
    
    If SEC(3).Value Or SEC(0).Value Then
        cString = cString & " GROUP BY FILE4_10.CODE, FILE4_10.DESCA, FACT.DESCA, FILE1_10.MOSM , FILE1_10.GRMODEL ORDER BY FILE4_10.CODE, FILE4_10.DESCA, FACT.DESCA, FILE1_10.MOSM, FILE1_10.GRMODEL "
    ElseIf SEC(2).Value Then
        cString = cString & " GROUP BY FILE4_10.CODE, FILE4_10.DESCA, FACT.DESCA, FILE0_82.DESCA , FILE1_10.GRMODEL ORDER BY FILE4_10.CODE, FILE4_10.DESCA, FACT.DESCA, FILE0_82.DESCA, FILE1_10.GRMODEL "
    ElseIf SEC(1).Value Then
        cString = cString & " GROUP BY FILE4_10.CODE, FILE4_10.DESCA, FACT.DESCA, FILE0_81.DESCA , FILE1_10.GRMODEL ORDER BY FILE4_10.CODE, FILE4_10.DESCA, FACT.DESCA, FILE0_81.DESCA, FILE1_10.GRMODEL "
    End If
    Set DATA17.Recordset = myCmd(cString, con)
End With
FIXGRID7
If grid7.Rows > 0 Then grid7.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FIXGRID7()
With grid7
    .Cols = 18
    .RowHeight(0) = 1000
    .RowHidden(1) = True
    .WordWrap = True
    .ColHidden(1) = True
    .TextMatrix(0, 0) = "«·„Ê—œ"
    .TextMatrix(0, 1) = "«·„’‰⁄"
        
    If SEC(3).Value Or SEC(0).Value Then
        .TextMatrix(0, 2) = "„Ê”„"
    ElseIf SEC(2).Value Then
        .TextMatrix(0, 2) = "›∆… ⁄„—Ì…"
    ElseIf SEC(1).Value Then
        .TextMatrix(0, 2) = "«·‰Ê⁄"
    End If
    
    .TextMatrix(0, 3) = "„Ã„Ê⁄… „ÊœÌ·"
    .TextMatrix(0, 4) = "»Ì«‰"
    
    .TextMatrix(0, 5) = "ﬂ„Ì… Ê«—œ"
    .TextMatrix(0, 6) = "ﬁÌ„… «·Ê«—œ"
    
    .TextMatrix(0, 7) = "ﬂ„Ì… ’«œ—"
    .TextMatrix(0, 8) = "ﬁÌ„… ’«œ—"
    
    .TextMatrix(0, 9) = "’«›Ï «·ﬂ„Ì…"
    .TextMatrix(0, 10) = "’«›Ï «·ﬁÌ„…"
    
    .TextMatrix(0, 11) = "‰”»… „»Ì⁄« "
    
    .TextMatrix(0, 12) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 13) = " ﬂ·›… „»Ì⁄« "
    
    .TextMatrix(0, 14) = "—’Ìœ ⁄œœ"
    .TextMatrix(0, 15) = "ﬁÌ„… «·—’Ìœ"
    
    .TextMatrix(0, 16) = "⁄œœ „ÊœÌ·« "
    .TextMatrix(0, 17) = "⁄„— Ã—Ê»"
    
    .ColHidden(6) = Not bOpt10
    .ColHidden(8) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    .ColHidden(13) = Not bOpt10
    .ColHidden(15) = Not bOpt10
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 0
    .ColWidth(2) = 1600
    .ColWidth(3) = 1100
    .ColWidth(4) = 2000
    
    .ColWidth(5) = 700
    .ColWidth(6) = 1100
    .ColWidth(7) = 700
    .ColWidth(8) = 1100
    .ColWidth(9) = 700
    .ColWidth(10) = 1100
    .ColWidth(11) = 900
    
    .ColWidth(12) = 700
    .ColWidth(13) = 1100
    
    .ColWidth(14) = 700
    .ColWidth(15) = 1100
    
    .ColWidth(16) = 550
    .ColWidth(17) = 550
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    For nCol = 5 To 9 Step 2
        .ColFormat(nCol) = "#0"
        .ColFormat(nCol + 1) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        .ColDataType(nCol + 1) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol + 1) = Format(.TextMatrix(nRow, nCol + 1), "#0.00")
        Next nRow
    Next nCol
    .ColFormat(12) = "#0"
    .ColFormat(13) = "#0.00"
    .ColDataType(12) = flexDTDouble
    .ColDataType(13) = flexDTDouble
    
    .ColFormat(14) = "#0"
    .ColFormat(15) = "#0.00"
    .ColDataType(14) = flexDTDouble
    .ColDataType(15) = flexDTDouble
    
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 9)) <> 0 Then .TextMatrix(nRow, 11) = Round((Val(.TextMatrix(nRow, 12)) / Val(.TextMatrix(nRow, 9))) * 100, 2)
    Next nRow
    .SubtotalPosition = flexSTAbove
    For I = 5 To .Cols - 2
        .Subtotal flexSTSum, -1, I, "#0", vbRed, vbYellow, True, "  "
    Next I
    If .Rows > 2 Then
        nRow = 2
        If Val(.TextMatrix(nRow, 9)) <> 0 Then .TextMatrix(nRow, 11) = Round((Val(.TextMatrix(nRow, 12)) / Val(.TextMatrix(nRow, 9))) * 100, 2)
    End If
    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub
Private Sub myload5()
Dim cString  As String
Dim cField As String, cWhere As String, cWhereField As String
With GRID5
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "([IN]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "(FILE1_11_ALL.[IN] * FILE1_11_ALL.COST ) ")
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "(FILE1_11_ALL.[IN] * FILE1_10.PRICE ) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "([OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "(FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST ) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "(FILE1_11_ALL.[OUT] * FILE1_10.PRICE ) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "([IN]-[OUT]) ")
    
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "(([IN]-[OUT]) * FILE1_11_ALL.COST ) ")
    cField = cField & "," & _
            myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )", "(([IN]-[OUT]) * FILE1_10.PRICE ) ")

    cField = cField & "," & _
            myiif("type = '6' ", " [OUT] ")
    
    cField = cField & "," & _
            myiif("type = '6' ", " ([OUT]* FILE1_11_ALL.COST) ")
    
    cField = cField & "," & _
            myiif("type = '6' ", " (FILE1_11_ALL.TOTAL ) ")
    
    cField = cField & "," & _
            myiif("type = '6' ", " FILE1_11_ALL.TOTAL -([OUT]* FILE1_11_ALL.COST) ")
                                
        
    cString = " SELECT  FILE4_10.DESCA AS supldesca, FACT.DESCA AS factdesca, FILE1_10.MOSM  ,  FILE1_10.MODELFACT0, FILE1_10.desca , STORE_ALL.DESCA  " & _
                cField & _
                " , file0_81.desca , file0_82.desca , file1_10SC.DESCA , FILE1_50.DESCA , GrModel ,MOSM2     " & _
                "  FROM    FILE1_10 INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE " & _
                " INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE" & _
                " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM" & _
                " INNER JOIN STORE_ALL ON STORE_ALL.CODE = FILE1_11_ALL.STORE " & _
                " LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.ModelSex" & _
                " LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.ModelAge" & _
                " LEFT JOIN FILE1_10SC ON FILE1_10SC.CODE = FILE1_10.[SECTION]" & _
                " LEFT JOIN FILE1_50 ON FILE1_50.CODE = FILE1_10.[GROUP]" & _
                " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE"
    
    If chkStoped.Value = 0 Then cWhere = "STORE_BR.ISSTOP = 0"
    If xMosm.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text) & " OR FILE1_10.MOSM  = '000'  )"
    Else
        If XMOSM2.text <> "" Then cWhere = cWhere & Tr(cWhere) & " ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M'  OR FILE1_10.MOSM  = '000' ) "
    End If
    
    If xmosm20.text <> "" Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MOSM2 = " & MyParn(xmosm20.text)
    
    If XSECTION.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xStore.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " [store] = " & MyParn(xStore.BoundText)
    
    If Not bOpt5 Then
        cWhere = cWhere & Tr(cWhere) & " [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    End If
    
    If XOutlet.Value = 1 Then cWhere = cWhere & Tr(cWhere) & " Outlet = 1 "
    
    If xage.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.MatchedWithList Then cWhere = cWhere & Tr(cWhere) & " FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)


    If IsDate(xDate1.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] >= " & DateSq(xDate1.Value)
    If IsDate(xDate2.Value) Then cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xDate2.Value)
    
    If xall(0).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 1 "
    ElseIf xall(1).Value Then
        cWhere = cWhere & Tr(cWhere) & " ISBRANCH = 0 "
    End If
    
    
    If cWhere <> "" Then
        cString = cString & " Where " & cWhere
    End If
    cString = cString & " GROUP BY FILE4_10.CODE, FILE4_10.DESCA,STORE_ALL.DESCA,STORE_ALL.CODE , FACT.DESCA, FILE1_10.MOSM, FILE1_10.MODELFACT0, FILE1_10.desca , FILE1_10.MODEL  , FILE1_10.desca , FILE1_10.MODEL   , file0_81.desca , file0_82.desca , file1_10SC.DESCA , FILE1_50.DESCA , GrModel ,MOSM2"
    cString = cString & " ORDER BY FILE4_10.CODE, FILE4_10.DESCA, FACT.DESCA, FILE1_10.MOSM, FILE1_10.MODELFACT0 ,STORE_ALL.CODE "
    
    If Check1.Value = 1 Then
        textquery.text = cString
    Else
        Set DATA15.Recordset = myRecordSet(cString, con)
    End If
       
End With
FIXGRID5
If GRID5.Rows > 0 Then GRID5.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub


Sub FIXGRID5()
With GRID5
    .Cols = 28
    .RowHeight(0) = 1000
    .RowHidden(1) = True
    .WordWrap = True
    
    .TextMatrix(0, 0) = "«·„Ê—œ"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .TextMatrix(0, 2) = "„Ê”„"
    .TextMatrix(0, 3) = "„ÊœÌ·"
    .TextMatrix(0, 4) = "«·’‰›"
    .TextMatrix(0, 5) = "«·›—⁄"
    
    .TextMatrix(0, 6) = "ﬂ„Ì… Ê«—œ"
    .TextMatrix(0, 7) = "‰ﬂ·›… «·Ê«—œ"
    .TextMatrix(0, 8) = " ﬁÌ„ «·Ê«—œ „” Â·ﬂ"
    
    .TextMatrix(0, 9) = "ﬂ„Ì… ’«œ—"
    .TextMatrix(0, 10) = " ﬂ·›… ’«œ—"
    .TextMatrix(0, 11) = " ﬁÌ„ ’«œ— „” Â·ﬂ"
    
    .TextMatrix(0, 12) = "’«›Ï «·ﬂ„Ì…"
    .TextMatrix(0, 13) = "’«›Ï ‰ﬂ·›… «·ﬁÌ„…"
    .TextMatrix(0, 14) = "’«›Ï „” Â·ﬂ "
    
    .TextMatrix(0, 15) = "ﬂ„Ì… „»Ì⁄« "
    .TextMatrix(0, 16) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 17) = "ﬁÌ„… „»Ì⁄« "
    .TextMatrix(0, 18) = "—»Õ „»Ì⁄« "
    
    .TextMatrix(0, 19) = "«·—’Ìœ"
    .TextMatrix(0, 20) = " ﬂ·›… «·—’Ìœ"
    .TextMatrix(0, 21) = " ﬁÌ„ —’Ìœ »«·„” Â·ﬂ"
    
    .ColHidden(19) = True
    .ColHidden(20) = True
    .ColHidden(21) = True
    
    
    .TextMatrix(0, 22) = "«·‰Ê⁄"
    .TextMatrix(0, 23) = "«·›∆… «·⁄„—Ì…"
    .TextMatrix(0, 24) = "«·ﬁ”„"
    .TextMatrix(0, 25) = "„Ã„Ê⁄… «·„ÊœÌ·« "
    .TextMatrix(0, 26) = "„Ã„Ê⁄… „ÊœÌ·"
    .TextMatrix(0, 27) = "„Ê”„ œ«Œ·Ï"
    
    
'    .ColHidden(6) = Not bOpt10
'    .ColHidden(8) = Not bOpt10
'    .ColHidden(10) = Not bOpt10
'    .ColHidden(12) = Not bOpt10
'    .ColHidden(14) = Not bOpt10
'    .ColHidden(16) = Not bOpt10
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1200
    .ColWidth(2) = 700
    .ColWidth(3) = 1500
    .ColWidth(4) = 3000
    .ColWidth(5) = 2000
    
    .ColWidth(6) = 900
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    
    .ColWidth(9) = 900
    .ColWidth(10) = 1200
    .ColWidth(11) = 1200
    
    .ColWidth(12) = 900
    .ColWidth(13) = 1200
    .ColWidth(14) = 1200
    
    .ColWidth(15) = 900
    .ColWidth(16) = 1200
    .ColWidth(17) = 1200
    .ColWidth(18) = 1200
    .ColWidth(19) = 1200
    .ColWidth(20) = 1200
    
    
    .ColWidth(21) = 1200
    .ColWidth(22) = 1200
    .ColWidth(23) = 1200
    .ColWidth(24) = 1200
    .ColWidth(25) = 1200
    .ColWidth(26) = 1200
    .ColWidth(27) = 1200
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    For nCol = 6 To 20
        .ColFormat(nCol) = "#0"
        .ColDataType(nCol) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol) = Format(.TextMatrix(nRow, nCol), "#0")
        Next nRow
    Next nCol
        
    For nCol = 6 To 20
        .ColFormat(nCol) = "#0"
        .ColFormat(nCol + 1) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        .ColDataType(nCol + 1) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol + 1) = Format(.TextMatrix(nRow, nCol + 1), "#0.00")
        Next nRow
    Next nCol
        
    .SubtotalPosition = flexSTAbove
    For I = 6 To 20
        .Subtotal flexSTSum, -1, I, "#0", vbRed, vbYellow, True, "  "
    Next I
    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub

Private Sub textquery_GotFocus()
textquery.SelStart = 0
textquery.SelLength = Len(textquery.text)
End Sub

