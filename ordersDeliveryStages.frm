VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form online_Stage_main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   4950
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   45
      Width           =   2670
      Begin Threed.SSCommand cmdUpdateStage 
         Height          =   510
         Left            =   45
         TabIndex        =   36
         Top             =   135
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
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
         Picture         =   "ordersDeliveryStages.frx":0000
         Caption         =   " ⁄œÌ· Õ«·… «·ÿ·»Ì« "
         TagVariant      =   "«Œ «— «·„ﬁ«Ê·"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   7650
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   90
      Width           =   4920
      Begin Threed.SSCommand cmdStage 
         Height          =   465
         Left            =   45
         TabIndex        =   32
         Top             =   135
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   820
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Œ «— „—Õ·… «· ÃÂÌ“"
         TagVariant      =   "«Œ «— «·„—Õ·…"
         ButtonStyle     =   3
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   7650
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   765
      Width           =   4920
      Begin VB.OptionButton optType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„€·ﬁ…"
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
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   225
         Width           =   735
      End
      Begin VB.OptionButton optType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„› ÊÕ…"
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
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton optType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„·€«…"
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
         Index           =   4
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   225
         Width           =   690
      End
      Begin VB.OptionButton optType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Index           =   0
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   225
         Width           =   690
      End
      Begin VB.OptionButton optType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   225
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1410
      Left            =   12600
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   7710
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   5175
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1410
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1260
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   1590
      End
      Begin VB.TextBox XDOC_NO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   5175
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   585
         Width           =   1410
      End
      Begin VB.TextBox xphone 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   945
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo xpay 
         Height          =   330
         Left            =   180
         TabIndex        =   3
         Top             =   585
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xMan 
         Height          =   330
         Left            =   180
         TabIndex        =   5
         Top             =   945
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ"
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
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   180
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï  «—ÌŒ"
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
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ ÿ·»Ì…"
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
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   585
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·”œ«œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   630
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «· ·Ì›Ê‰"
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
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   945
         Width           =   945
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„‰œÊ»"
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
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   945
         Width           =   705
      End
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   24
      Top             =   10440
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   315
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   7305
      Begin Threed.SSCommand cmdCSV 
         Height          =   510
         Left            =   4410
         TabIndex        =   14
         Top             =   135
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
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
         Picture         =   "ordersDeliveryStages.frx":25D3
         Caption         =   "CSV ”Õ»"
         TagVariant      =   "«Œ «— «·„ﬁ«Ê·"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   510
         Left            =   2790
         TabIndex        =   15
         Top             =   135
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
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
         Picture         =   "ordersDeliveryStages.frx":4BA6
         Caption         =   " ’œÌ— Excel"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   510
         Left            =   1395
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   135
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   900
         _Version        =   196610
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
         Picture         =   "ordersDeliveryStages.frx":7010
         Caption         =   " ÿ»«⁄…"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "ordersDeliveryStages.frx":9386
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   510
         Left            =   5940
         TabIndex        =   33
         Top             =   135
         Width           =   1320
         _ExtentX        =   2328
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
         Picture         =   "ordersDeliveryStages.frx":B509
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "ordersDeliveryStages.frx":E3BC
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   -315
      Top             =   -135
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7305
      Left            =   90
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1440
      Width           =   20220
      _cx             =   35666
      _cy             =   12885
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
      Cols            =   29
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   18
      Top             =   10635
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   741
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   19
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   714
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   1
         Left            =   4095
         TabIndex        =   20
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   2
         Left            =   8100
         TabIndex        =   21
         Top             =   45
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   3
         Left            =   12150
         TabIndex        =   22
         Top             =   45
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   4
         Left            =   16155
         TabIndex        =   23
         Top             =   45
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "Online_Stage_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearch As New Search3, lIsBoxOnline As Boolean
Dim cString As String
Dim CLIST As String
Dim oSearchMan As New Search_abd
Dim oSearchStage As New Search_rs
Dim cStr1 As String, cStr2 As String
Dim oSearchItem As New Search3
Private Sub cmdCSV_Click()
Set onlineCSVfrm.myForm = Me
onlineCSVfrm.Show 1
End Sub
Private Sub cmdExcel_Click()
ToFileExelNew grid1, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption)
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
myload
End Sub

Private Sub cmdStage_Click()
StagesOnlineLook Me, oSearchStage, , , IIf(cmdStage.Tag = "", "", "ﬂ· «·„—«Õ·")
End Sub

Private Sub cmdUpdateStage_Click()
Set onlineCSVfrm.myForm = Me
onlineCSVfrm.bNoItems = True
onlineCSVfrm.Show 1
End Sub
Private Sub Form_Load()
Dim db As New clsDb
Set xPay.RowSource = db.myRs("SELECT Payment_type " & _
                             "FROM file6_90h " & _
                             "where payment_type is not null  " & _
                             " GROUP BY Payment_type")

xPay.ListField = "Payment_type"
xPay.BoundColumn = "Payment_type"
Set xman.RowSource = db.myRs("SELECT CODE,DESCA FROM FILE6_25 WHERE FILE6_25.BRANCH IN (SELECT FILE0_40.BRANCH FROM FILE0_40 WHERE FILE0_40.online =  1) AND FILE6_25.ISSTOP = 0 ORDER BY CODE")
xman.ListField = "Desca"
xman.BoundColumn = "Code"
Set db = Nothing

fixGrd
End Sub
Public Sub myload()
Dim i As Double
Dim cString  As String, cStr2 As String
Dim cWhere As String
Dim cSubTotal As String
With grid1
    cString = " SELECT DOC_NO," & _
              " FORMAT(DATE,'yyyy/M/d')," & _
              " NAME," & _
              " PHONE," & _
              "Shipping_City," & _
              "STREET, " & _
              "FILE6_90H.TOTAL_ITEM" & "," & _
              "DISCOUNT, " & _
              "FILE6_90H.TOTAL_ITEM - FILE6_90H.DISCOUNT ," & _
              "DISCOUNT_CODE," & _
              "SHIPPING, " & _
              "FILE6_90H.TOTAL_ITEM + SHIPPING - FILE6_90H.DISCOUNT ," & _
              "Payment_type," & _
              "STORE," & _
              "STAGES_CODES.DESCA ," & _
              "FILE6_25.DESCA," & _
              "SEND_USER," & _
              "FORMAT(SEND_TIME,'yyyy/M/d')," & _
              "NOTES_ORDER ," & _
              "SUBSTRING(SALES_DOC,7,20) ," & _
              "FORMAT(SALES_DATE,'yyyy/M/d')," & _
              "SHIP_NO," & _
              "FORMAT(SHIP_DATE,'yyyy/M/d'),"
cString = cString & _
              " NOTES," & _
              " FORMAT(DelOrder_Date,'yyyy/M/d')," & _
              " FILE6_90H.MAN," & _
              " FILE6_90H.STAGE," & _
              " FILE6_90H.TAGS," & _
              " FILE6_90H.CLOSED" & _
              " FROM FILE6_90H " & _
              " LEFT JOIN FILE6_25 ON FILE6_90H.MAN = FILE6_25.CODE" & _
              " INNER JOIN STAGES_CODES ON FILE6_90H.STAGE = STAGES_CODES.CODE"
              
    If optType(1).Value Then
        cWhere = cWhere & Tr(cWhere) & "FILE6_90H.DOC_NO IN (SELECT ORDER_NO FROM vw_online_orders_open)"
    ElseIf optType(2).Value Then
        cWhere = cWhere & Tr(cWhere) & "(FILE6_90H.DOC_NO IN (SELECT ORDER_NO FROM vw_online_orders_closed) OR FILE6_90H.CANCELED = 1)"
    ElseIf optType(3).Value Then
        cWhere = cWhere & Tr(cWhere) & "FILE6_90H.DOC_NO IN (SELECT ORDER_NO FROM vw_online_orders_closed)"
    ElseIf optType(4).Value Then
        cWhere = cWhere & Tr(cWhere) & "FILE6_90H.CANCELED = 1"
    End If
    
    If cmdStage.Tag <> "" Then
        cWhere = cWhere & Tr(cWhere) & "(FILE6_90H.STAGE = " & cmdStage.Tag & _
                 " OR (FILE6_90H.DOC_NO IN(SELECT ORDER_NO FROM FILE6_90BH WHERE FILE6_90BH.STAGE = " & cmdStage.Tag & ")))"
    End If
    
    If xdoc_no.text <> "" Then
        cWhere = cWhere & Tr(cWhere) & " [DOC_NO] = " & MyParn(xdoc_no.text)
    End If
    
    If xphone.text <> "" Then
        cWhere = cWhere & Tr(cWhere) & " [phone] = " & MyParn(xphone.text)
    End If
    
    If xPay.BoundText <> "" Then
        cWhere = cWhere & Tr(cWhere) & " [Payment_type] = " & MyParn(xPay.text)
    End If
    
    
    If xman.MatchedWithList Then
        cWhere = cWhere & Tr(cWhere) & "MAN = " & MyParn(xman.BoundText)
    End If
    
    If IsDate(xDate1.text) Then
        cWhere = cWhere & Tr(cWhere) & " [DATE] >= " & DateSq(xDate1.text)
    End If
    
    If IsDate(xdate2.text) Then
        cWhere = cWhere & Tr(cWhere) & " [DATE] <= " & DateSq(xdate2.text)
    End If
    
    'If xSales(0).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " SALES_DOC IS NULL"
    'If xSales(1).Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " SALES_DOC IS NOT NULL "
    
    'If RET_NOTSHIP.Value <> 0 Then cWhere = cWhere & Tr(cWhere) & " DelOrder_Date IS NOT NULL AND DelOrder_Date2 IS NULL "
    
    If cWhere <> "" Then cString = cString & " WHERE " & cWhere
    cString = cString & " ORDER BY DOC_NO "
    'On Error GoTo myerror
    Set grid1.DataSource = myRs(cString)
End With
fixGrd
grid1.Cell(flexcpAlignment, 0, 0, grid1.Rows - 1, grid1.Cols - 1) = 7
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub fixGrd()
With grid1
    .RowHeight(0) = 600
    .WordWrap = True
    .FrozenCols = 3
    
    .TextMatrix(0, 0) = "—ﬁ„"
    .TextMatrix(0, 1) = "«· «—ÌÕ"
    .TextMatrix(0, 2) = "«·≈”„"
    .TextMatrix(0, 3) = " ·Ì›Ê‰"
    .TextMatrix(0, 4) = "«·„œÌ‰…"
    .TextMatrix(0, 5) = "⁄‰Ê«‰"
    
    .TextMatrix(0, 6) = "≈Ã„«·Ï «’‰«›"
    .TextMatrix(0, 7) = "Œ’„"
    .TextMatrix(0, 8) = "«·≈Ã„«·Ï"
    .TextMatrix(0, 9) = "ﬂÊœ Œ’„"
    .TextMatrix(0, 10) = "«·‘Õ‰"
    .TextMatrix(0, 11) = "«·≈Ã„«·Ì"
    
    .TextMatrix(0, 12) = "ÿ—Ìﬁ… «·œ›⁄"
    .TextMatrix(0, 13) = "«·›—⁄"
    .TextMatrix(0, 14) = "«·„—Õ·…"
    .TextMatrix(0, 15) = "«·„‰œÊ»"
    .TextMatrix(0, 16) = "«·„—”·"
    .TextMatrix(0, 17) = "SEND TIME "
    .TextMatrix(0, 18) = "„·«ÕŸ«  ÿ·»Ì…"
    .TextMatrix(0, 19) = "»Ê‰ „»Ì⁄« "
    .TextMatrix(0, 20) = " «—ÌŒ «·»Ì⁄"
    
    .TextMatrix(0, 21) = "»Ê·Ì’… ‘Õ‰"
    .TextMatrix(0, 22) = " «—ÌŒ «·‘Õ‰"
    .TextMatrix(0, 23) = "„·ÕÊŸ… «·«·€«¡"
    .TextMatrix(0, 24) = " «—ÌŒ «·«·€«¡"
    .TextMatrix(0, 25) = "«·„‰œÊ»"
    .TextMatrix(0, 26) = "«·„—Õ·…"
    .TextMatrix(0, 27) = "«· √ﬂÌœ"
    .TextMatrix(0, 28) = "„€·ﬁ…"
        
    .ColHidden(9) = True
    .ColHidden(13) = True
    '.ColHidden(12) = True
    .ColHidden(15) = True
    .ColHidden(16) = True
    .ColHidden(21) = True
    .ColHidden(25) = True
    .ColHidden(26) = True
    .ColHidden(28) = True
    
    '.TextMatrix(0, 22) = "”»» «·≈·€«¡"
    '.TextMatrix(0, 23) = "«·€«¡ «·ÿ·»"
    
    '.TextMatrix(0, 24) = " «—ÌŒ «” ·«„ «·„— Ã⁄"
    '.TextMatrix(0, 25) = "„·«ÕŸ«  «·—∆Ì”Ì"
    '.TextMatrix(0, 26) = "„·«ÕŸ«  «·›—Ê⁄"
    '.TextMatrix(0, 27) = "«·„‰œÊ»"

    
    .ColDataType(1) = flexDTDate
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ColDataType(9) = flexDTDouble
    .ColDataType(10) = flexDTDouble
    
                    
    .ColWidth(0) = 800
    .ColWidth(1) = 1300
    .ColWidth(2) = 2000
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 4000
    .ColWidth(6) = 1000
    .ColWidth(7) = 900
    .ColWidth(8) = 1300
    .ColWidth(9) = 800
    .ColWidth(10) = 900
    .ColWidth(11) = 1400
    .ColWidth(12) = 1800
    .ColWidth(13) = 1200
    .ColWidth(14) = 1500
    .ColWidth(15) = 1800
    
    .ColWidth(18) = 1000
    .ColWidth(17) = 2000
    .ColWidth(19) = 1300
    .ColWidth(20) = 1250
    .ColWidth(22) = 1300
    .ColWidth(23) = 1300
    .ColWidth(24) = 1300
    
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
    .Subtotal flexSTSum, -1, 8, "#", vbYellow, , True, ""
    .Subtotal flexSTSum, -1, 9, "#", vbYellow, , True, ""
    .Subtotal flexSTSum, -1, 10, "#", vbYellow, , True, ""
    .Subtotal flexSTSum, -1, 11, "#", vbYellow, , True, ""
    
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'closeCon con
    'If cBranch <> "00" Then closeCon con_MyShop
    SaveText Me, , Array(xDate1.Name, xdate2.Name)
End Sub
Private Sub grid1_DblClick()
If grid1.Row < 1 Or grid1.Row = grid1.Rows - 1 Then Exit Sub

Dim db As New clsDb
Dim nCount As Integer
If Not db.rsValue("select count(*) " & _
              " from file6_90Bh" & _
              " Where order_no = " & MyParn(grid1.TextMatrix(grid1.Row, 0)) & _
              " and closed = 0", nCount) Then
    Set db = Nothing
    Exit Sub
End If

Set db = Nothing

If grid1.ValueMatrix(grid1.Row, 28) = 0 Or nCount = 0 Or cBranch = "00" Then
    orders_online_items.sDoc_no = grid1.TextMatrix(grid1.Row, 0)
    orders_online_items.Show 1
Else
    orders_online_invoices.bedit = True
    orders_online_invoices.sOrder_no = grid1.TextMatrix(grid1.Row, 0)
    orders_online_invoices.sStore = sStoreOnline
    orders_online_invoices.sStage = cmdStage.Tag
    orders_online_invoices.bOpenOnly = True

    Set orders_online_invoices.myForm = Me
    orders_online_invoices.Show 1
End If
End Sub
Private Sub grid1_GotFocus()
'grid1_EnterCell
End Sub
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = cmdStage.Name Then
    If oSearchStage.grid1.TextMatrix(oSearchStage.grid1.Row, 0) = "" Then
        cmdStage.Tag = ""
        cmdStage.Caption = cmdStage.TagVariant
    Else
        cmdStage.Tag = oSearchStage.grid1.TextMatrix(oSearchStage.grid1.Row, 0)
        cmdStage.Caption = oSearchStage.grid1.TextMatrix(oSearchStage.grid1.Row, 1)
    End If
    Unload oSearchStage
    myload
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub SSCommand1_Click()
End Sub

Private Sub xPhone_GotFocus()
myGotFocus xphone
End Sub
Private Sub XPHONE_LostFocus()
myLostFocus xphone
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xdoc_no
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xdoc_no
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xdate2
myValidDate xdate2
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xPay_GotFocus()
myGotFocus xPay
End Sub
Private Sub xPay_LostFocus()
myLostFocus xPay
If Not xPay.MatchedWithList Then xPay.BoundText = ""
End Sub
Private Sub XSTORE_GotFocus()
myGotFocus XSTORE
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus XSTORE
If Not XSTORE.MatchedWithList Then XSTORE.BoundText = ""
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grid1.col = 12 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
    If grid1.col = 13 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
    If grid1.col = 19 + 1 Then Exit Sub
    If grid1.col = 21 + 1 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
    If grid1.col = 24 + 1 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    If col = 12 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
    If col = 13 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
    If col = 19 + 1 Then Exit Sub
    If col = 21 + 1 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
    If col = 24 + 1 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
    CellPos KeyCode, Row, col
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If col < 2 Then
    grid1.col = 2
ElseIf col > 2 And col < 5 Then
    grid1.col = grid1.col + 1
ElseIf col < 12 Then
    grid1.col = 12
    grid1.ShowCell Row, grid1.col
ElseIf col < 13 Then
    grid1.col = 13
    grid1.ShowCell Row, grid1.col
'ElseIf Col < 19 And cBranch = "00" Then
'    grid1.Col = 19
'    grid1.ShowCell Row, grid1.Col
ElseIf col < 22 + 1 And cBranch = "00" Then
    grid1.col = 22 + 1
    grid1.ShowCell Row, grid1.col
ElseIf col < 23 + 1 And cBranch <> "00" Then
    grid1.col = 23 + 1
    grid1.Row = Row
    grid1.ShowCell Row, grid1.col
ElseIf col < 24 + 1 Then
    grid1.col = 24 + 1
    grid1.Row = Row
    grid1.ShowCell Row, grid1.col
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 2)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, col
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If col = 4 Then
    If Val(grid1.EditText) = 0 Then
        Cancel = True
        Exit Sub
    End If
    
    If Val(xTotal.Caption) > 0 And Val(grid1.EditText) <= 0 Then
        Cancel = True
        Exit Sub
    ElseIf Val(xTotal.Caption) < 0 And Val(grid1.EditText) >= 0 Then
        Cancel = True
        Exit Sub
    End If
ElseIf col = 19 Then
    If (Not IsDate(grid1.EditText)) And Trim(grid1.EditText) <> "" Then
        Cancel = True
    Else
        grid1.EditText = myFormat_p(grid1.EditText)
    End If
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.col
End If
End Sub
