VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form price_listfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   20355
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10860
   ScaleWidth      =   20355
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   0
      ScaleHeight     =   1080
      ScaleWidth      =   20355
      TabIndex        =   35
      Top             =   9315
      Width           =   20355
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   0
         Width           =   3525
         Begin Threed.SSCommand cmdFirst 
            Height          =   420
            Left            =   2610
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   135
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   741
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
            Picture         =   "price_lists.frx":0000
            Caption         =   "√Ê·"
            ButtonStyle     =   3
            PictureAlignment=   10
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "price_lists.frx":21A7
         End
         Begin Threed.SSCommand cmdPrevious 
            Height          =   420
            Left            =   1710
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   135
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   741
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
            Picture         =   "price_lists.frx":41EE
            Caption         =   "”«»ﬁ"
            ButtonStyle     =   3
            PictureAlignment=   10
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "price_lists.frx":62D9
         End
         Begin Threed.SSCommand cmdNext 
            Height          =   420
            Left            =   855
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   135
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   741
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
            Picture         =   "price_lists.frx":82D3
            Caption         =   "·«Õﬁ"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "price_lists.frx":A3E4
         End
         Begin Threed.SSCommand cmdLast 
            Height          =   420
            Left            =   45
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   135
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   741
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
            Picture         =   "price_lists.frx":C3DE
            Caption         =   "√ŒÌ—"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "price_lists.frx":E602
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   -45
         Width           =   6360
         Begin VB.TextBox txtModel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   90
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Tag             =   "1"
            Top             =   225
            Width           =   2040
         End
         Begin VB.TextBox txtBarcode 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3420
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Tag             =   "1"
            Top             =   225
            Width           =   2040
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "—ﬁ„ „ÊœÌ·"
            Height          =   285
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   225
            Width           =   825
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "»«—ﬂÊœ"
            Height          =   285
            Left            =   5535
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   270
            Width           =   645
         End
      End
      Begin Threed.SSCommand cmdClient 
         Height          =   555
         Left            =   10035
         TabIndex        =   41
         Top             =   90
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   979
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
         Caption         =   "«Œ «— «·⁄„Ì·"
         TagVariant      =   "«Œ «— «·⁄„Ì·"
         ButtonStyle     =   3
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   375
         Left            =   45
         TabIndex        =   47
         Top             =   675
         Visible         =   0   'False
         Width           =   14100
         _ExtentX        =   24871
         _ExtentY        =   661
         _Version        =   327682
         BorderStyle     =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   10530
      TabIndex        =   34
      Top             =   -90
      Width           =   4110
      Begin Threed.SSCommand cmdExcel 
         Height          =   510
         Left            =   45
         TabIndex        =   48
         Top             =   180
         Width           =   1905
         _ExtentX        =   3360
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
         Picture         =   "price_lists.frx":106D3
         Caption         =   " ÕÊÌ· «·Ì „·› "
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdCSV 
         Height          =   510
         Left            =   1935
         TabIndex        =   49
         Top             =   180
         Width           =   2130
         _ExtentX        =   3757
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
         Picture         =   "price_lists.frx":12B3D
         Caption         =   "”Õ» „‰ „·› CSV"
         TagVariant      =   "«Œ «— «·„ﬁ«Ê·"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
   End
   Begin VB.Frame Frame99 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5235
      Left            =   -4005
      TabIndex        =   9
      Top             =   11070
      Visible         =   0   'False
      Width           =   10725
      Begin VB.CommandButton cmd_showdoc 
         Caption         =   "⁄—÷  ›’Ì·Ï «·»Ê‰"
         Height          =   450
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   270
         Width           =   2310
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "„— Ã⁄ "
         Height          =   270
         Left            =   1035
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label xTotalQuant2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1035
         Width           =   1005
      End
      Begin VB.Label xusername_RET 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   2610
         TabIndex        =   14
         Top             =   990
         Width           =   1950
      End
      Begin VB.Label XUSERNAME_DISC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   2610
         TabIndex        =   13
         Top             =   1350
         Width           =   1950
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "»Ê‰ «·„»Ì⁄« "
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
         Left            =   675
         TabIndex        =   12
         Top             =   765
         Width           =   960
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„— Ã⁄"
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
         Left            =   3195
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame FRAME_CUR 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Index           =   4
      Left            =   6165
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   675
      Width           =   1275
      Begin Threed.SSCommand cmdSave 
         Height          =   555
         Left            =   45
         TabIndex        =   22
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "price_lists.frx":15110
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "price_lists.frx":17A35
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   510
         Left            =   45
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   720
         Width           =   1185
         _ExtentX        =   2090
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
         Picture         =   "price_lists.frx":1A289
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "price_lists.frx":1C3E9
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   14670
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   -90
      Width           =   5595
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   4500
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   180
         Width           =   1050
         _ExtentX        =   1852
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
         Picture         =   "price_lists.frx":1E6D6
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "price_lists.frx":20AA1
      End
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   3375
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "price_lists.frx":22B4A
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "price_lists.frx":24B52
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   2250
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "price_lists.frx":26B09
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "price_lists.frx":292A5
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   180
         Width           =   1050
         _ExtentX        =   1852
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
         Picture         =   "price_lists.frx":2B739
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   510
         Left            =   1125
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "price_lists.frx":2DA5C
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "price_lists.frx":2FDD2
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   -4770
      Top             =   8145
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   28
      Top             =   10395
      Width           =   20355
      _ExtentX        =   35904
      _ExtentY        =   820
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
         TabIndex        =   29
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
         TabIndex        =   30
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
         TabIndex        =   31
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
         TabIndex        =   32
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
         TabIndex        =   33
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   7470
      TabIndex        =   5
      Top             =   630
      Width           =   12750
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10485
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   0
         TabStop         =   0   'False
         Tag             =   "2"
         Top             =   180
         Width           =   1185
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   10485
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   1185
      End
      Begin VB.TextBox xCode2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   5355
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2460
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6975
         MaxLength       =   200
         TabIndex        =   2
         Top             =   900
         Width           =   4695
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6975
         TabIndex        =   27
         Top             =   540
         Width           =   3480
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "«·⁄„Ì·"
         Height          =   270
         Left            =   11745
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   585
         Width           =   450
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ „” ‰œ"
         Height          =   285
         Left            =   11745
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   750
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "⁄„·«¡ ﬂ—Ê  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   2430
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·»Ì«‰"
         Height          =   270
         Left            =   11745
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   990
         Width           =   420
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7215
      Left            =   90
      TabIndex        =   3
      Top             =   2025
      Width           =   20220
      _cx             =   35666
      _cy             =   12726
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
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
   Begin MSComDlg.CommonDialog Common1 
      Left            =   270
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "price_listfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bAct As Boolean
Dim cFilter As String
Dim bIg As Boolean
Dim bStopCell As Boolean
Public sDoc_no As String
Dim cFilterLook As String
Dim nRound As Long
Dim CardTable As ADODB.RecordSet
Dim oSearchItem As New Search_abd, oSearchDoc  As New Search_abd
Dim osearchClient As New Search_abd, oSearchClientLook As New Search_abd
Dim bEdit As Boolean
Dim bEditRecord As Boolean
Dim con As New ADODB.Connection
Dim formMode
Const LoadMode = 0, DefineMode = 1
Sub ItemsLookup()
ItemsLookupAll Me, oSearchItem
End Sub
Function myreplace(Optional Row As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(aInsert, "CODE", addstring(xCode.text))
aInsert = AddFlag(aInsert, "[NOTES]", addstring(xNotes.text))
On Error GoTo myerror
con.BeginTrans
If xDoc_No.Tag = DefineMode Then
    xDoc_No.text = NewFlag("FILE6_53H", "DOC_NO", con)
    aInsert = AddFlag(aInsert, "DOC_NO", xDoc_No.text)
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(cusername))
    con.Execute addInsert(aInsert, "FILE6_53H")
Else
    aInsert = AddFlag(aInsert, "[USERNAME2]", addstring(cusername))
    con.Execute addUpdate(aInsert, "FILE6_53H", "DOC_NO = " & addstring(xDoc_No.text))
End If
If Row > 0 Then myreplaceGrd Row
con.CommitTrans
myreplace = True
Exit Function
myerror:
'prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
'On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    Dim bNew As Boolean
    bNew = grid1.Row = grid1.Rows - 1
    grid1.TextMatrix(grid1.Row, 1) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    CellPos 13, grid1.Row, grid1.Col
    GrdDesc grid1.TextMatrix(grid1.Row, 1), grid1.Row
    Grid1_AfterEdit grid1.Row, grid1.Col
ElseIf ActiveControl.Name = cmdInform.Name Then
    xDoc_No.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
    myUndo
ElseIf ActiveControl.Name = xCode.Name Then
    xCode.text = oSearchClientLook.grid1.TextMatrix(oSearchClientLook.grid1.Row, 0)
    oSearchClientLook.Hide
    SendKeys "{tab}"
ElseIf ActiveControl.Name = cmdClient.Name Then
    If osearchClient.grid1.TextMatrix(osearchClient.grid1.Row, 0) = "" Then
        cmdClient.Tag = ""
        cmdClient.Caption = cmdClient.TagVariant
    Else
        cmdClient.Tag = osearchClient.grid1.TextMatrix(osearchClient.grid1.Row, 0)
        cmdClient.Caption = osearchClient.grid1.TextMatrix(osearchClient.grid1.Row, 1)
    End If
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then myDefine
    End If
    osearchClient.Hide
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub chkDay_Click()
myUndo
End Sub
Private Sub cmdClient_Click()
CLIENTLOOKUP Me, osearchClient, , , IIf(cmdClient.Tag = "", "", "ﬂ· «·⁄„·«¡")
End Sub

Private Sub cmdCSV_Click()
'If grid1.Rows > 2 Then
'    MsgBox " ÊÃœ «’‰«› »«·„” ‰œ"
'    Exit Sub
'End If
getData
End Sub

Private Sub CmdDel_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) <> vbOK Then
    Exit Sub
End If

con.BeginTrans
con.Execute "Delete  From FILE6_53 where Doc_No = " & xDoc_No.text
con.Execute "Delete  From FILE6_53H where Doc_No = " & xDoc_No.text
con.CommitTrans
    
If Not openCardTable(tbMode.tbPrevious, xDoc_No.text) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExcel_Click()
Me.MousePointer = 11
Dim aRow As Variant
    If grid1.Rows > 1 Then
        aSub = AddFlag(Empty, "row", 1)
        aSub = AddFlag(aSub, "bold", True)
        aSub = AddFlag(aSub, "word_wrap", False)
        aSub = AddFlag(aSub, "back_color", 19)
        aRow = AddFlag(aRow, aSub)
    End If
    ToFileExelNew grid1, , , aRow, , 0.9, , , , , , Me, Array("⁄—÷ ”⁄— —ﬁ„ : " & xDoc_No.text, "«·⁄„Ì· : " & xCodeDesca.Caption, "«· «—ÌŒ : " & xDate.text)

Me.MousePointer = 0
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdInform_Click()
CardLookup
End Sub

Private Sub cmdMan_Click()

End Sub

Private Sub CmdNewInv_Click()
myDefine
On Error Resume Next
xCode.SetFocus
Err.Clear
End Sub
Private Sub cmdSave_Click()
If myreplace Then
    MsgBox " „ «·Õ›Ÿ »‰Ã«Õ"
End If
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub CmdUndo_Click()
myUndo
End Sub
Private Sub Form_Activate()
If Not bAct Then
    If xDoc_No.Tag = LoadMode Then
        On Error Resume Next
        grid1.SetFocus
        CellPos 13, grid1.Rows - 2, grid1.Cols - 1
    Else
       xCode.SetFocus
    End If
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then KeyAscii = 0
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        SendKeys "{TAB}"
        KeyCode = 0
    End If
ElseIf KeyCode = 116 And cmdSave.Enabled Then
    cmdSave_Click
End If
End Sub
Private Sub Form_Load()
'On Error GoTo myerror
bEdit = True

OpenCon con
nRound = 0

If Not openCardTable Then myDefine
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Form_Resize()
Dim nHeight As Long
nHeight = Me.Height - (grid1.Top + 2000)
If nHeight > 3000 Then
    grid1.Height = nHeight
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

CardTable.Close
Set CardTable = Nothing

closeCon con

Set sales_ordersfrm = Nothing
Err.Clear
End Sub

Public Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'On Error GoTo myerror
Dim bNew As Boolean
With grid1
If Not myValid(True) Then
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
    myLoadGrd
    If Row < grid1.Rows - 1 Then
        grid1.Select Row, Col
    Else
        CellPos 13, grid1.Rows - 2, grid1.Cols - 1
    End If
    Exit Sub
End If


If Not validRow(Row) Then
    Exit Sub
End If

If Row = grid1.Rows - 1 Then
    myAddItem
    bNew = True
End If

If myreplace(Row) Then
    If xDoc_No.Tag = DefineMode Then
        openCardTable tbMode.tbFind, xDoc_No.text
    ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        grid1.TextMatrix(Row, grid1.Cols - 1) = get_Id(con)
    End If
    If bNew Then
        grid1.Select grid1.Rows - 1, 1
        grid1.ShowCell grid1.Rows - 1, 1
    End If
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub grid1_EnterCell()
With grid1
    If Not bEditRecord Then
        .Editable = flexEDNone
    ElseIf .Col = 1 Or .Col = 11 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub grid1_GotFocus()
grid1_EnterCell
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'If grid1.Col = 12 And grid1.TextMatrix(grid1.Row, grid1.Col) = "" Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    'If Col = 12 And grid1.TextMatrix(Row, Col) = "" Then Exit Sub
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim sBarCode As String
Dim nSal As Double
Dim nRet As Double
Dim cPassword As String, aUser As Variant
With grid1
    If Col = 9 Or Col = 10 Then
        If Not IsNumeric(.EditText) Then
            Cancel = True
            Exit Sub
        End If
    ElseIf Col = 1 Then
        If Not GrdDesc(.EditText, Row) Then
            Cancel = True
        Else
            If .EditText <> grid1.TextMatrix(Row, 1) Then
                .EditText = grid1.TextMatrix(Row, 1)
            End If
        End If
    End If
End With
End Sub
Private Sub cmdPrint_Click()
If chkModel.Value = 0 Then
    doprint
Else
    doPrintModel
End If
End Sub
Private Sub txtBarcode_Change()
SearchGrd grid1, LCase(Trim(txtBarcode.text)), 1, True
End Sub

Private Sub txtModel_Change()
SearchGrd grid1, LCase(Trim(txtModel.text)), 5, False
End Sub
Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CLIENTLOOKUP2 Me, oSearchClientLook, "isstop = 0"
End Sub
Private Sub xDate_Validate(Cancel As Boolean)
If Not IsDate(xDate.text) Then Cancel = True
End Sub
Private Function myValid(Optional bIgMsg As Boolean = False, Optional bIgMan As Boolean = False) As Boolean
If Trim(xCode.text) = "" Then
    MsgBox "·« ÌÊÃœ ﬂÊœ ⁄„Ì· «Ê ﬂÊœ „Ê—œ"
    Exit Function
End If

If xCodeDesca.Caption = "" Then
    If Not bIgMsg Then MsgBox "·« ÌÊÃœ ﬂÊœ ⁄„Ì·"
    Exit Function
End If
myValid = True
End Function
Private Sub myload()
xDoc_No.text = CardTable!doc_no
xNotes.text = CardTable!notes & ""
xCode.text = CardTable!code & ""
xCodeDesca.Caption = CardTable!ClientDesca & ""
panel1(0).Caption = CardTable!UserName & ""

Handlecontrols LoadMode

myLoadGrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
Exit Sub
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub myDefine()
xDoc_No.text = NewFlag("FILE6_53H", "DOC_NO", con)

xCode.text = ""
xCodeDesca.Caption = ""


xNotes.text = ""
panel1(1).Caption = GetComputerName
panel1(2).Caption = cusername

grid1.Rows = 1
myAddItem

fixGrd

Handlecontrols DefineMode

CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub Handlecontrols(nMode)
bEditRecord = bEdit
cmdNewInv.Enabled = bEdit And nMode = LoadMode

xDoc_No.Tag = nMode


cmdSave.Enabled = bEditRecord
cmddel.Enabled = (bEditRecord Or bDelRecord) And nMode = LoadMode

Dim nRecord As Long, nRecords As Long
retRecords xDoc_No.text, nRecords, nRecord

cmdNext.Enabled = nRecord < nRecords And nRecords <> 0 And nMode = LoadMode
cmdPrevious.Enabled = nRecord <> 1 And nRecords <> 0 And nMode = LoadMode
cmdLast.Enabled = nRecord < nRecords And nRecords <> 0 And nMode = LoadMode
cmdFirst.Enabled = nRecord <> 1 And nRecords <> 0 And nMode = LoadMode
cmdPrint.Enabled = nMode = LoadMode


If nMode = LoadMode Then
    panel1(0).Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
    panel1(0).Caption = "”Ã· ÃœÌœ (" & (nRecords + 1) & ")"
End If

End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If Not bEditRecord Then
    Exit Sub
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
ElseIf KeyCode = 112 And grid1.Editable = flexEDKbdMouse Then
    ItemsLook Me, oSearchItem
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 2 Then
    RemoveItem (grid1.Row)
    grid1.Select grid1.Rows - 1, 1
    grid1.ShowCell grid1.Rows - 1, 1
    MakeSerial grid1.Row
End If
End Sub
Private Function GrdDesc(pItem As String, ByRef Row As Long, Optional ByRef sMsg As String) As Boolean
With grid1
If Trim(pItem) = "" Or Len(pItem) < 3 Or Len(pItem) > 30 Then
    sMsg = "ﬂÊœ €Ì— ’«·Õ"
    Exit Function
End If

Dim cm As New ADODB.command
Set cm = cmd("dbo.sp_item", con, adStoredProc, AddFlag(Empty, "BARCODE_FIND", pItem))
cm.Execute

If IsNull(cm.Parameters("@ITEM").Value) Then
    sMsg = "ﬂÊœ €Ì— ’«·Õ"
    Exit Function
End If

'If Not cm.Parameters("@ISNOITEM").Value Then
'    If fnBalance(cm.Parameters("@ITEM").Value, con, xStore.BoundText, xDate.text) <= 0 Then
'        Inform "·« ÌÊÃœ —’Ìœ «·’‰›"
'    End If
'End If

grid1.TextMatrix(Row, 1) = cm.Parameters("@ITEM").Value
grid1.TextMatrix(Row, 2) = cm.Parameters("@MOSM").Value
grid1.TextMatrix(Row, 3) = cm.Parameters("@FACT_DESCA").Value
grid1.TextMatrix(Row, 4) = cm.Parameters("@SUPP").Value & ""
grid1.TextMatrix(Row, 5) = cm.Parameters("@MODELFACT0").Value
grid1.TextMatrix(Row, 6) = cm.Parameters("@DESCA").Value
grid1.TextMatrix(Row, 7) = cm.Parameters("@SCAL").Value
grid1.TextMatrix(Row, 8) = cm.Parameters("@COLOR").Value
grid1.TextMatrix(Row, 9) = cm.Parameters("@PRICE").Value
grid1.TextMatrix(Row, 10) = cm.Parameters("@COST2").Value
GrdDesc = True
End With
End Function
Private Sub CardLookup(Optional pFilter As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)
Set Generalarray(0) = Me

Generalarray(1) = "SELECT TOP 1000 " & _
                  " FILE6_53H.DOC_NO ," & _
                  " FILE3_10.DESCA," & _
                  " FILE6_53H.NOTES" & _
                  " FROM FILE6_53H" & _
                  " INNER JOIN FILE3_10 ON FILE6_53H.CODE = FILE3_10.CODE"

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & " WHERE " & cFilter
End If

Generalarray(2) = "Order by doc_no DESC"
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "«·⁄„Ì·"
listarray(0, 1) = "(%%FILE3_10.DESCA%%)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1200

GrdArray(1, 0) = "«·⁄„Ì·"
GrdArray(1, 1) = 6500

GrdArray(2, 0) = "≈”„ «·ﬁ«∆„…"
GrdArray(2, 1) = 3500

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub
Private Sub xPrinted_Click()
If bIg Then Exit Sub
con.Execute "update FILE6_53h set printed = " & xPrinted.Value & ",closed = " & xPrinted.Value & " from FILE6_53h where doc_no = " & MyParn(xDoc_No.text)
MsgBox " „ › Õ «·›« Ê—… »‰Ã«Õ"
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then
        myDefine
    End If
End If
End Sub
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Sub fixGrd()
With grid1
'                 0         1          2         3         4          5             6          7          8               9          10              11
.FormatString = "„.|" & "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "”⁄— „” Â·ﬂ|" & "”⁄— Ã„·…|" & "”⁄—|"
.RowHeight(0) = 600
.WordWrap = True


.ColWidth(0) = 600
.ColWidth(1) = 1000
.ColWidth(2) = 0
.ColWidth(3) = 2000
.ColWidth(4) = 0
.ColWidth(5) = 2000
.ColWidth(6) = 5000
.ColWidth(7) = 1200
.ColWidth(8) = 1600
.ColWidth(9) = 1200
.ColWidth(10) = 1200
.ColWidth(11) = 1200

.ColHidden(.Cols - 1) = True
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
For i = 0 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Function RemoveItem(Row As Long) As Boolean
If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel) <> vbOK Then
    Exit Function
End If

If grid1.TextMatrix(Row, grid1.Cols - 1) <> "" Then
    On Error GoTo myerror
    con.Execute "Delete  From FILE6_53 where id = " & grid1.TextMatrix(Row, grid1.Cols - 1)
End If
grid1.RemoveItem Row
RemoveItem = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function myreplaceGrd(Row) As Boolean
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "ITEM", addstring(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "PRICE", .ValueMatrix(i, 11))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.text))
            con.Execute addInsert(aInsert, "FILE6_53")
        Else
            con.Execute addUpdate(aInsert, "FILE6_53", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub myLoadGrd()
Dim strSql As String

strSql = "SELECT FILE6_53.ITEM," & _
         "FILE1_10.MOSM," & _
         "FACT.DESCA," & _
         "FILE1_10.SUPP," & _
         "FILE1_10.MODELFACT0," & _
         "FILE1_10.DESCA," & _
         "FILE1_10.SCAL," & _
         "FILE1_10.COLOR," & _
         "FILE1_10.PRICE," & _
         "FILE1_10.COST2," & _
         "FILE6_53.PRICE," & _
         "FILE6_53.ID " & _
         "FROM FILE6_53 " & _
         "INNER JOIN FILE1_10 " & _
         " ON FILE6_53.ITEM = FILE1_10.ITEM" & _
         " LEFT JOIN FACT ON FACT.CODE = FILE1_10.FACT" & _
         " WHERE FILE6_53.DOC_NO = " & xDoc_No.text & _
         " ORDER BY FILE6_53.ID"
Set grid1.DataSource = cmd(strSql, con).Execute
myAddItem
fixGrd
End Sub
Private Function mysave() As Boolean

If Not myValid Then Exit Function

If Not myreplace Then Exit Function

Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
                    
End Function
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = True) As Boolean
With grid1
If Not IsNumeric(.TextMatrix(Row, 1)) Then Exit Function
If .ValueMatrix(Row, 11) = 0 Then Exit Function
End With
validRow = True
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf Col = 1 And grid1.TextMatrix(Row, Col) = "" Then
ElseIf Col < 11 Then
    grid1.Col = 11
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 11)
    grid1.ShowCell grid1.Row, 1
End If
End Sub
Private Sub myAddItem()
grid1.AddItem ""
MakeSerial
End Sub
Private Sub myUndo()
If xDoc_No.Tag = DefineMode Then
    If Not openCardTable Then
        CmdNewInv_Click
    End If
Else
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdNext_Click()
'If Not openCardTable(tbMode.tbNext, xdoc_no_flag.Caption) Then
If Not openCardTable(tbMode.tbNext, xDoc_No.text) Then
    If Not openCardTable(tbMode.tblast) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdPrevious_Click()
'If Not openCardTable(tbMode.tbPrevious, xdoc_no_flag.Caption) Then
If Not openCardTable(tbMode.tbPrevious, xDoc_No.text) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdFirst_Click()
If Not openCardTable(tbMode.tbFirst) Then
    myDefine
End If
End Sub
Private Sub CmdLast_Click()
If Not openCardTable Then
    myDefine
End If
End Sub
Private Function validOpen() As Boolean
If cBranch <> "00" Then
    MsgBox "„»Ì⁄«  ›—⁄ —∆Ì”Ì ›ﬁÿ"
    Exit Function
End If
validOpen = True
End Function
Private Function retRecords(pDoc_no, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New ADODB.RecordSet
If pDoc_no <> "" Then
    cString = "SELECT Count(FILE6_53H.DOC_NO) AS records,COUNT(CASE WHEN FILE6_53H.DOC_NO <= " & MyParn(pDoc_no) & " THEN 1 END) AS record"
Else
    cString = "SELECT Count(FILE6_53H.DOC_NO) AS records,0 as record"
End If

cString = cString & " FROM FILE6_53H " & _
                    " INNER JOIN FILE3_10 ON FILE6_53H.CODE = FILE3_10.CODE"

If cFilter <> "" Then
    cString = cString & " WHERE " & cFilter
End If

Set loctable = cmd(cString, con).Execute
If Not loctable.EOF Then
    nRecords = loctable!RECORDS
    nRecord = Val(loctable!Record & "")
End If
End Function
Private Function retFilter()
If cmdClient.Tag <> "" Then
    retFilter = retFilter & " AND FILE6_53H.CODE = " & MyParn(cmdClient.Tag)
End If
End Function
Private Sub xdate_filter_GotFocus()
myGotFocus xdate_filter
End Sub
Private Sub xdate_filter_LostFocus()
myLostFocus xdate_filter
myValidDate xdate_filter
End Sub
Private Sub xTotal_GotFocus()
myGotFocus xTotal
End Sub
Private Sub xDiscount_GotFocus()
myGotFocus xDiscount
End Sub
Private Sub xRate_GotFocus()
myGotFocus xRate
End Sub
Private Sub xInv_No_GotFocus()
myGotFocus xinv_no
End Sub
Private Sub xInv_No_LostFocus()
myLostFocus xinv_no
End Sub
Private Sub xcard_disc_GotFocus()
myGotFocus xcard_disc
End Sub
Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xdate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub xcode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xcode_LostFocus()
myLostFocus xCode
xCodeDesca.Caption = ""
If Trim(xCode.text) = "" Then Exit Sub
xCode.text = RetZero(xCode.text, 4)
If xCode.text = "0000" Then Exit Sub

Dim loctable As New ADODB.RecordSet
Set loctable = cmd("SELECT * FROM FILE3_10 WHERE CODE = " & MyParn(xCode.text), con).Execute
If loctable.EOF Then Exit Sub
xCodeDesca.Caption = loctable!DESCA & ""
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_No
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No
If Not ValidNum(xDoc_No.text) Then
     If xDoc_No.Tag = LoadMode Then
        myDefine
    Else
        xDoc_No.text = ""
    End If
Else
    If (Not (CardTable.EOF)) And xDoc_No.Tag = LoadMode Then
        If CardTable!doc_no = xDoc_No.text Then
            Exit Sub
        End If
    End If
    
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
End If
End Sub
Private Sub XSTORE_GotFocus()
myGotFocus xStore
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus xStore
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub doprint()
Dim temptable As New ADODB.RecordSet
Dim sourcetable As New ADODB.RecordSet
Dim lPrice As Boolean
Dim nFBal As Double
Dim nPay As Double
Dim nBal As Double

Dim nWidth As Long, nHeight As Long
Dim imageTable As New ADODB.RecordSet

imgxPrint.Images.Clear

If xLogo.MatchedWithList Then
    Set imageTable = cmd("select * from logo_codes where code = " & IIf(xLogo.MatchedWithList, xLogo.BoundText, 1), con).Execute
    If Not imageTable.EOF Then
        If Not IsNull(imageTable!Image) Then
            Set picPrint.DataSource = imageTable
            picPrint.DataField = "image"
            
            GetNewSize picPrint.Width, picPrint.Height, 500, nHeight
            'aRet = retDimImg(picPrint.Width, picPrint.Height, 600, 200)
            
            imgxPrint.Import.FromMemoryFile (imageTable!Image)
            imgxPrint.Effects.Resize 500, nHeight
        End If
    End If
End If

Dim loctable As New ADODB.RecordSet, cString As String
cString = "SELECT FILE6_53H.DOC_NO," & _
          " FILE6_53H.DATE," & _
          " FILE3_10.DESCA AS CODE_DESCA," & _
          " FILE6_53.ITEM," & _
          " FILE6_53.QUANT," & _
          " FILE6_53.PRICE," & _
          " FILE1_10.DESCA AS ITEM_DESCA," & _
          " FILE6_53.TOTAL," & _
          " file1_10.modelfact," & _
          " FACT.DESCA AS FACTDESCA," & _
          " FILE1_10.COLOR," & _
          " FILE1_10.C_COLOR," & _
          " FILE1_10.SCAL," & _
          " FILE1_10.MODEL" & _
          " FROM FILE6_53 " & _
          " INNER JOIN FILE6_53H ON FILE6_53.DOC_NO = FILE6_53H.DOC_NO" & _
          " INNER JOIN FILE1_10 ON FILE6_53.ITEM = FILE1_10.ITEM" & _
          " INNER JOIN FILE3_10 ON FILE6_53H.CODE = FILE3_10.CODE" & _
          " LEFT JOIN FACT ON FACT.CODE = FILE1_10.[FACT]" & _
          " WHERE  FILE6_53.DOC_NO = " & MyParn(xDoc_No.text)
Set loctable = cmd(cString, con).Execute

Dim con_image As New ADODB.Connection
openConDoc con_image

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
Do While Not loctable.EOF
    temptable.AddNew
    
    cString = " select top 1 * " & _
             " from MODEL_IMAGES where model = " & MyParn(loctable!MODEL) & _
             " AND (C_COLOR = " & loctable!c_Color & _
             " OR C_COLOR IS NULL)" & _
             " ORDER BY C_COLOR,ID"
     
    If chkImage.Value = 1 Then
        Set imageTable = cmd(cString, con_image).Execute
        ImgxPrint2.Images.Clear
        If Not imageTable.EOF Then
            If Not IsNull(imageTable!Image) Then
                ' Skip picPrint binding completely if it causes errors
                ImgxPrint2.Import.FromMemoryFile (imageTable!Image)
                
                ' Get dimensions directly from the third-party control instead of picPrint
                ' (Note: check your control's documentation for exact property names like .Width or .ImageWidth)
                aRet = retDimImg(ImgxPrint2.Width, ImgxPrint2.Height, 300, 300)
                
                ImgxPrint2.Effects.Resize retFlag(aRet, "width"), retFlag(aRet, "height")
                SaveImageToRs ImgxPrint2.Picture, temptable, "picture2"
            End If
        End If
    End If
     
'    If chkImage.Value = 1 Then
'        Set imageTable = cmd(cString, con_image).Execute
'        ImgxPrint2.Images.Clear
'        If Not imageTable.EOF Then
'            If Not IsNull(imageTable!Image) Then
'                Set picPrint.Picture = LoadPicture("")
'                Set picPrint.DataSource = imageTable
'                picPrint.DataField = "image"
'
'                'GetNewSize picPrint.Width, picPrint.Height, nWidth, 200
'
'                aRet = retDimImg(picPrint.Width, picPrint.Height, 300, 300)
'
'                ImgxPrint2.Import.FromMemoryFile (imageTable!Image)
'                ImgxPrint2.Effects.Resize retFlag(aRet, "width"), retFlag(aRet, "height")
'                'ImgxPrint2.Effects.Resize nWidth, 200
'
'                SaveImageToRs ImgxPrint2.Picture, temptable, "picture2"
'            End If
'        End If
'    End If
    
    If imgxPrint.Images.Count > 0 Then
        SaveImageToRs imgxPrint.Picture, temptable, "picture1"
    End If
    
    temptable!STR6 = loctable!doc_no
    temptable!Date1 = myFormat_p(loctable!Date)
    temptable!STR7 = loctable!CODE_desca
    
    temptable!str2 = loctable!FACTDESCA
    temptable!str1 = loctable!Item
    temptable!str13 = TurnValue(DelZero(loctable!modelfact))
    temptable!str3 = loctable!ITEM_DESCA
    temptable!STR4 = loctable!color
    temptable!STR5 = loctable!SCAL
    temptable!str16 = loctable!MODEL
    temptable!str10 = "⁄—÷ ”⁄—"
    
    temptable!VAL3 = loctable!Quant
    temptable!VAL1 = loctable!price
    temptable!VAL4 = loctable!TOTAL
    temptable!VAL14 = i
    temptable.Update
    loctable.MoveNext
Loop

closeCon con_image

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
REPORT1.Destination = crptToWindow
REPORT1.ReportFileName = App.Path & "\Reports\Print_SALES" & IIf(chkQuant.Value = 1, "_Quant", "") & IIf(chkImage.Value = 1, "_images", "") & ".rpt"
REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub doPrintModel()
Dim temptable As New ADODB.RecordSet
Dim sourcetable As New ADODB.RecordSet
Dim lPrice As Boolean
Dim nFBal As Double
Dim nPay As Double
Dim nBal As Double

Dim nWidth As Long, nHeight As Long
Dim imageTable As New ADODB.RecordSet

imgxPrint.Images.Clear
If xLogo.MatchedWithList Then
    Set imageTable = cmd("select * from logo_codes where code = " & IIf(xLogo.MatchedWithList, xLogo.BoundText, 1), con).Execute
    If Not imageTable.EOF Then
        If Not IsNull(imageTable!Image) Then
            Set picPrint.DataSource = imageTable
            picPrint.DataField = "image"
            
            GetNewSize picPrint.Width, picPrint.Height, 500, nHeight
            imgxPrint.Import.FromMemoryFile (imageTable!Image)
            imgxPrint.Effects.Resize 500, nHeight
        End If
    End If
End If

Dim loctable As New ADODB.RecordSet
Dim cString As String
cString = "SELECT FILE1_10.MODEL," & _
            "FILE1_10.MODELFACT," & _
            "FILE1_10.DESCA AS ITEM_DESCA," & _
            "MAX(FILE6_53.PRICE) AS PRICE," & _
            "FILE6_53H.DOC_NO," & _
            "FILE6_53H.[DATE]," & _
            "FACT.DESCA AS FACT_DESCA," & _
            "FILE3_10.DESCA AS CODE_desca," & _
            " (SELECT  STRING_AGG(COLOR, CHAR(13)) AS COLOR FROM (SELECT DISTINCT  COLOR FROM FILE1_10 F where model = FILE1_10.MODEL) FC) AS COLOR," & _
            " (SELECT  STRING_AGG(SCAL, CHAR(13)) AS SCAL FROM (SELECT DISTINCT  SCAL FROM FILE1_10 F where model = FILE1_10.MODEL) FS) AS SCAL" & _
            " From FILE6_53 " & _
            " INNER JOIN FILE6_53H ON FILE6_53.DOC_NO = FILE6_53H.DOC_NO" & _
            " INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE6_53.ITEM" & _
            " INNER JOIN FILE3_10 ON FILE6_53H.CODE = FILE3_10.CODE" & _
            " LEFT JOIN FACT ON FACT.CODE = FILE1_10.[FACT]" & _
            " WHERE FILE6_53H.DOC_NO = " & MyParn(xDoc_No.text)
 
 cString = cString & _
            " GROUP BY FILE6_53H.DOC_NO," & _
             "FILE1_10.MODEL," & _
             "FILE1_10.MODELFACT," & _
             "FILE1_10.DESCA," & _
             "FILE6_53.DOC_NO," & _
             "FILE6_53H.[DATE]," & _
             "FILE3_10.DESCA," & _
             "FACT.DESCA" & _
            " ORDER BY FILE1_10.MODEL"

Set loctable = myRs(cString, con)

Dim con_image As New ADODB.Connection
openConDoc con_image

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
Do While Not loctable.EOF
    temptable.AddNew
    
    cString = "SELECT TOP 1 * " & _
             " FROM MODEL_IMAGES WHERE MODEL = " & MyParn(loctable!MODEL) & _
             " ORDER BY C_COLOR,ID"
     
    If chkImage.Value = 1 Then
        Set imageTable = cmd(cString, con_image).Execute
        ImgxPrint2.Images.Clear
        If Not imageTable.EOF Then
            If Not IsNull(imageTable!Image) Then
                ' Skip picPrint binding completely if it causes errors
                ImgxPrint2.Import.FromMemoryFile (imageTable!Image)
                
                ' Get dimensions directly from the third-party control instead of picPrint
                ' (Note: check your control's documentation for exact property names like .Width or .ImageWidth)
                aRet = retDimImg(ImgxPrint2.Width, ImgxPrint2.Height, 300, 300)
                
                ImgxPrint2.Effects.Resize retFlag(aRet, "width"), retFlag(aRet, "height")
                SaveImageToRs ImgxPrint2.Picture, temptable, "picture2"
            End If
        End If
    End If
     
'        Set imageTable = cmd(cString, con_image).Execute
'        ImgxPrint2.Images.Clear
'        If Not imageTable.EOF Then
'            If Not IsNull(imageTable!Image) Then
'                Set picPrint.Picture = LoadPicture("")
'                Set picPrint.DataSource = imageTable
'                picPrint.DataField = "image"
'
'                'GetNewSize picPrint.Width, picPrint.Height, nWidth, 200
'
'                aRet = retDimImg(picPrint.Width, picPrint.Height, 300, 300)
'
'                ImgxPrint2.Import.FromMemoryFile (imageTable!Image)
'                ImgxPrint2.Effects.Resize retFlag(aRet, "width"), retFlag(aRet, "height")
'                'ImgxPrint2.Effects.Resize nWidth, 200
'
'                SaveImageToRs ImgxPrint2.Picture, temptable, "picture2"
'            End If
'        End If
    
    If imgxPrint.Images.Count > 0 Then
        SaveImageToRs imgxPrint.Picture, temptable, "picture1"
    End If
    
    temptable!STR6 = loctable!doc_no
    temptable!Date1 = myFormat_p(loctable!Date)
    temptable!STR7 = loctable!CODE_desca
    temptable!VAL1 = loctable!price
    
    temptable!str2 = loctable!fact_Desca
    temptable!MEMO1 = loctable!color
    temptable!MEMO2 = loctable!SCAL
    temptable!str16 = loctable!MODEL
    temptable!str13 = TurnValue(DelZero(loctable!modelfact))
    temptable!str3 = loctable!ITEM_DESCA
    temptable!str10 = "⁄—÷ ”⁄—"
    temptable!VAL14 = i
    temptable.Update
    loctable.MoveNext
Loop

closeCon con_image

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If

contemp.BeginTrans
contemp.CommitTrans
REPORT1.Destination = crptToWindow
REPORT1.ReportFileName = App.Path & "\Reports\Print_SALES_MODEL.rpt"
REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Public Function IsFormOpen(ByVal FormName As String) As Boolean
    Dim frm As Form
    For Each frm In Forms
        If UCase(frm.Name) = UCase(FormName) Then
            IsFormOpen = True
            Exit Function
        End If
    Next
    IsFormOpen = False
End Function
Private Function openCardTable(Optional pMode As Integer = tbMode.tblast, Optional pDoc_no As String = "", Optional pDoc_no_Filter As String = "") As Boolean
Dim cString As String
Dim cWhere As String
Dim cOrder As String
Dim nRecords As Long, nRecord As Long

If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pDoc_no) = "" And Trim(pDoc_no_Filter) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass


'On Error GoTo myerror
cFilter = retFilter

Set CardTable = New ADODB.RecordSet
cString = "SELECT TOP 1 FILE6_53H.*," & _
          " FILE3_10.DESCA AS CLIENTDESCA " & _
          " FROM FILE6_53H " & _
          " INNER JOIN FILE3_10 ON FILE3_10.CODE = FILE6_53H.code"
  
If pMode = tbMode.tbFirst Then
    'cOrder = "Order by FILE6_53H.DOC_NO_FLAG"
    cOrder = "Order by FILE6_53H.DOC_NO"
ElseIf pMode = tbMode.tblast Then
    'cOrder = "Order by FILE6_53H.DOC_NO_FLAG DESC"
    cOrder = "Order by FILE6_53H.DOC_NO DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "FILE6_53H.Doc_no = " & pDoc_no
ElseIf pMode = tbMode.tbPrevious Then
    'cWhere = "FILE6_53H.DOC_NO_FLAG < " & MyParn(pDoc_no)
    'cOrder = "Order by FILE6_53H.DOC_NO_FLAG DESC"
    
    cWhere = "FILE6_53H.DOC_NO < " & pDoc_no
    cOrder = "Order by FILE6_53H.DOC_NO DESC"
ElseIf pMode = tbMode.tbNext Then
    'cWhere = "FILE6_53H.DOC_NO_FLAG > " & MyParn(pDoc_no)
    cWhere = "FILE6_53H.DOC_NO > " & MyParn(pDoc_no)
    cOrder = "Order by FILE6_53H.DOC_NO"
End If

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

If cFilter <> "" Then
    cString = cString & IIf(cWhere = "", " WHERE ", " AND ") & cFilter
End If

cString = cString & " " & cOrder

Set CardTable = cmd(cString, con).Execute

If (Not CardTable.EOF) Then
    myload
    openCardTable = True
End If
Me.MousePointer = vbNormal
Exit Function
myerror:
Me.MousePointer = vbNormal
MsgBox Err.Description
Err.Clear
End Function
Private Function getData()
Dim cFileName As String
Dim sb As New ChilkatStringBuilder

Common1.InitDir = ""
Common1.FileName = ""
Common1.Filter = "CSV (*.CSV*)|*.CSV*"
Common1.ShowOpen

cFileName = Common1.FileName

If cFileName = "" Then Exit Function

Dim cSv As New ChilkatCsv
Dim i As Long

nAccess = sb.LoadFile(cFileName, "auto")

If nAccess = 0 Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰  Õ„Ì· «·„·›"
    Exit Function
End If


cSv.HasColumnNames = 1
nAccess = cSv.LoadFromString(sb.GetAsString)

If cSv.NumRows < 1 Then Exit Function

If myloadCSV(cSv) Then
    Inform " „  «÷«›… «·«’‰«› »‰Ã«Õ"
    myLoadGrd
End If
End Function
Private Function myloadCSV(ByRef cSv As ChilkatCsv) As Boolean
Dim i As Long
Dim sCaption As String
Dim ship_no As String


sCaption = Me.Caption
prog1.Visible = True

Dim sb As New ChilkatStringBuilder
Dim nAffect As Long
For i = 0 To cSv.NumRows - 1
    If Trim(cSv.GetCell(i, 0)) <> "" Then
        sb.Append MyParn(cSv.GetCell(i, 0)) & ","
    End If
Next

sb.Shorten 1
If sb.length = 0 Then Exit Function

Dim strSql As String
Dim ItemTable As ADODB.RecordSet

strSql = "SELECT * " & _
         " FROM FILE1_10 " & _
         " WHERE ITEM IN(" & sb.GetAsString & ")"

Dim db As New clsDb
Set ItemTable = db.myRs(strSql): If ItemTable Is Nothing Then GoTo CleanUp

                    

Dim nquant As Double
prog1.Visible = True
For i = 1 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
        
    sitem = Trim(cSv.GetCell(i, 0))
    nPrice = Val(cSv.GetCell(i, 1))
                        
    If sitem <> "" And nPrice > 0 Then
         ItemTable.Find "ITEM = " & MyParn(sitem), , adSearchForward, adBookmarkFirst
         If Not ItemTable.EOF Then
            If nPrice > 0 Then
                strSql = "update file6_53 " & _
                              "set file6_53.price = " & nPrice & _
                              " where doc_no = " & MyParn(xDoc_No.text) & _
                              " and item  = " & MyParn(sitem)
                
                If Not db.Execute(strSql, nAffect) Then GoTo CleanUp
                
                If nAffect = 0 Then
                    strSql = "INSERT INTO FILE6_53(DOC_NO,ITEM,PRICE)" & _
                             "VALUES(" & xDoc_No.text & "," & _
                             addstring(sitem) & "," & _
                             nPrice & ")"
                    If Not db.Execute(strSql) Then GoTo CleanUp
                End If
            End If
        End If
    End If
Next
myloadCSV = True
CleanUp:
prog1.Visible = False
Me.Caption = sCaption
Set db = Nothing
End Function
