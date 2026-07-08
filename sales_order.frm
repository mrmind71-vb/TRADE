VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form sales_ordersfrm 
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9855
      TabIndex        =   78
      Top             =   -90
      Width           =   4785
      Begin VB.CheckBox chkModel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " „ÊœÌ·« "
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   180
         TabIndex        =   83
         Top             =   270
         Width           =   1140
      End
      Begin VB.CheckBox chkQuant 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "”⁄— ›ﬁÿ"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1890
         TabIndex        =   80
         Top             =   270
         Width           =   1140
      End
      Begin VB.CheckBox chkImage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ÿ»«⁄… ’Ê—"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3420
         TabIndex        =   79
         Top             =   270
         Width           =   1230
      End
   End
   Begin VB.Frame fmPrice 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4770
      TabIndex        =   71
      Top             =   1260
      Width           =   3075
      Begin VB.OptionButton optPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "”⁄— Ã„·…"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   270
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.OptionButton optPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "”⁄— „” Â·ﬂ"
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   270
         Width           =   1275
      End
   End
   Begin VB.CheckBox XISINVOICE 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "›« Ê—…"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   -1170
      RightToLeft     =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   1575
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   0
      ScaleHeight     =   1080
      ScaleWidth      =   20355
      TabIndex        =   49
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
         TabIndex        =   50
         Top             =   0
         Width           =   3525
         Begin Threed.SSCommand cmdFirst 
            Height          =   420
            Left            =   2610
            TabIndex        =   51
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
            Picture         =   "sales_order.frx":0000
            Caption         =   "√Ê·"
            ButtonStyle     =   3
            PictureAlignment=   10
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "sales_order.frx":21A7
         End
         Begin Threed.SSCommand cmdPrevious 
            Height          =   420
            Left            =   1710
            TabIndex        =   52
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
            Picture         =   "sales_order.frx":41EE
            Caption         =   "”«»ﬁ"
            ButtonStyle     =   3
            PictureAlignment=   10
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "sales_order.frx":62D9
         End
         Begin Threed.SSCommand cmdNext 
            Height          =   420
            Left            =   855
            TabIndex        =   53
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
            Picture         =   "sales_order.frx":82D3
            Caption         =   "·«Õﬁ"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "sales_order.frx":A3E4
         End
         Begin Threed.SSCommand cmdLast 
            Height          =   420
            Left            =   45
            TabIndex        =   54
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
            Picture         =   "sales_order.frx":C3DE
            Caption         =   "√ŒÌ—"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "sales_order.frx":E602
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   -45
         Width           =   6360
         Begin VB.TextBox txtModel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   90
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   63
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
            TabIndex        =   62
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
            TabIndex        =   65
            Top             =   225
            Width           =   825
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "»«—ﬂÊœ"
            Height          =   285
            Left            =   5535
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   270
            Width           =   645
         End
      End
      Begin VB.Frame Frame16 
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
         Height          =   1050
         Left            =   14175
         TabIndex        =   55
         Top             =   0
         Width           =   6135
         Begin VB.Label xTotalItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   3015
            TabIndex        =   59
            Top             =   540
            Width           =   1905
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„… «’‰«›"
            Height          =   270
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   540
            Width           =   900
         End
         Begin VB.Label xTotalQuant 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   3015
            TabIndex        =   57
            Top             =   180
            Width           =   1905
         End
         Begin VB.Label lblQuant 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ ﬁÿ⁄"
            Height          =   270
            Left            =   4995
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   225
            Width           =   630
         End
      End
      Begin Threed.SSCommand cmdClient 
         Height          =   555
         Left            =   10035
         TabIndex        =   60
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
         TabIndex        =   70
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
      Left            =   7875
      TabIndex        =   47
      Top             =   1260
      Width           =   1905
      Begin Threed.SSCommand cmdExcel 
         Height          =   510
         Left            =   45
         TabIndex        =   48
         Top             =   180
         Width           =   1815
         _ExtentX        =   3201
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
         Picture         =   "sales_order.frx":106D3
         Caption         =   " ÕÊÌ· «·Ì „·› "
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         ShapeSize       =   1
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
      TabIndex        =   11
      Top             =   11070
      Visible         =   0   'False
      Width           =   10725
      Begin VB.CommandButton cmd_showdoc 
         Caption         =   "⁄—÷  ›’Ì·Ï «·»Ê‰"
         Height          =   450
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   12
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "Frame14"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4875
      Left            =   -9180
      TabIndex        =   17
      Top             =   3555
      Visible         =   0   'False
      Width           =   9285
      Begin VB.CheckBox xIs_collect 
         Caption         =   "Check1"
         Height          =   270
         Left            =   360
         TabIndex        =   68
         Top             =   765
         Width           =   1365
      End
      Begin VB.TextBox xinv_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1935
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3825
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox XISONEST 
         Caption         =   "Check1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2745
         TabIndex        =   30
         Top             =   1170
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CheckBox XISRETS 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3555
         TabIndex        =   29
         Top             =   2070
         Width           =   195
      End
      Begin VB.PictureBox pctBarCode 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         Left            =   1035
         ScaleHeight     =   118
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   151
         TabIndex        =   19
         Top             =   4095
         Visible         =   0   'False
         Width           =   2325
      End
      Begin MSAdodcLib.Adodc data1 
         Height          =   330
         Left            =   1125
         Top             =   3285
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
      Begin MSAdodcLib.Adodc DATA3 
         Height          =   330
         Left            =   2610
         Top             =   2655
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
      Begin MSAdodcLib.Adodc data4 
         Height          =   330
         Left            =   1035
         Top             =   3060
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
      Begin MSAdodcLib.Adodc data5 
         Height          =   330
         Left            =   2295
         Top             =   2880
         Visible         =   0   'False
         Width           =   2340
         _ExtentX        =   4128
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
      Begin MSAdodcLib.Adodc data12 
         Height          =   330
         Left            =   5040
         Top             =   1080
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
         Left            =   3015
         Top             =   2880
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
      Begin MSAdodcLib.Adodc data8 
         Height          =   330
         Left            =   4680
         Top             =   2070
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
      Begin Threed.SSCommand SSCommand1 
         Height          =   330
         Left            =   3330
         TabIndex        =   69
         Top             =   405
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " ÕÊÌ· «·Ì ›« Ê—… „Ã„⁄…"
         TagVariant      =   "«Œ «— «·„Œ“‰"
         ButtonStyle     =   3
      End
      Begin VB.Label xdoc_no_flag 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2655
         TabIndex        =   66
         Top             =   1485
         Width           =   1545
      End
      Begin VB.Label xdoc_disc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   3735
         TabIndex        =   35
         Top             =   1035
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label XBRANCH 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„” Œœ„"
         Height          =   285
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   555
      End
      Begin VB.Label xlate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   360
         TabIndex        =   33
         Top             =   2385
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„Œ“‰"
         Height          =   270
         Index           =   0
         Left            =   2805
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   3285
         Width           =   540
      End
      Begin VB.Label XTOTAL2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   555
         Left            =   4950
         TabIndex        =   31
         Top             =   855
         Width           =   2535
      End
      Begin VB.Label xcard_doc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   1440
         TabIndex        =   18
         Top             =   405
         Visible         =   0   'False
         Width           =   675
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
      Left            =   9810
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   675
      Width           =   1275
      Begin Threed.SSCommand cmdSave 
         Height          =   555
         Left            =   45
         TabIndex        =   27
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
         Picture         =   "sales_order.frx":12B3D
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_order.frx":15462
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   510
         Left            =   45
         TabIndex        =   28
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
         Picture         =   "sales_order.frx":17CB6
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_order.frx":19E16
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   14670
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   -90
      Width           =   5595
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   4500
         TabIndex        =   21
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
         Picture         =   "sales_order.frx":1C103
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_order.frx":1E4CE
      End
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   3375
         TabIndex        =   22
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
         Picture         =   "sales_order.frx":20577
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_order.frx":2257F
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   2250
         TabIndex        =   23
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
         Picture         =   "sales_order.frx":24536
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_order.frx":26CD2
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   24
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
         Picture         =   "sales_order.frx":29166
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   510
         Left            =   1125
         TabIndex        =   25
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
         Picture         =   "sales_order.frx":2B489
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "sales_order.frx":2D7FF
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   1440
      Top             =   675
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowProgressCtls=   0   'False
      WindowShowPrintSetupBtn=   -1  'True
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   -765
      Top             =   -225
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
   Begin Crystal.CrystalReport Report_T 
      Left            =   3015
      Top             =   3195
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc DATA13 
      Height          =   330
      Left            =   5535
      Top             =   360
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   41
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
         TabIndex        =   42
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
         TabIndex        =   43
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
         TabIndex        =   44
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
         TabIndex        =   45
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
         TabIndex        =   46
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
   Begin MSComDlg.CommonDialog Common1 
      Left            =   585
      Top             =   990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   11115
      TabIndex        =   6
      Top             =   630
      Width           =   9150
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
         Left            =   6975
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
         Left            =   6975
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
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
         TabIndex        =   5
         Top             =   2460
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3465
         MaxLength       =   200
         TabIndex        =   3
         Top             =   900
         Width           =   4695
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   1635
      End
      Begin MSDataListLib.DataCombo xLogo 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   90
         TabIndex        =   81
         Top             =   945
         Width           =   2355
         _ExtentX        =   4154
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·‘⁄«—"
         Height          =   270
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   990
         Width           =   480
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3465
         TabIndex        =   40
         Top             =   540
         Width           =   3480
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "«·⁄„Ì·"
         Height          =   270
         Left            =   8235
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   585
         Width           =   450
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ „” ‰œ"
         Height          =   285
         Left            =   8190
         RightToLeft     =   -1  'True
         TabIndex        =   7
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
         TabIndex        =   10
         Top             =   2430
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "„·«ÕŸ« "
         Height          =   270
         Left            =   8235
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   945
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
         Height          =   270
         Index           =   0
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   510
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7215
      Left            =   90
      TabIndex        =   4
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
      Cols            =   16
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
   Begin VB.Frame frPrinter 
      Height          =   4335
      Left            =   270
      RightToLeft     =   -1  'True
      TabIndex        =   74
      Top             =   3105
      Visible         =   0   'False
      Width           =   12075
      Begin VB.PictureBox picPrint 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   945
         Left            =   225
         ScaleHeight     =   945
         ScaleWidth      =   1125
         TabIndex        =   75
         Top             =   360
         Width           =   1125
      End
      Begin ImgXCtrl6.ImgXCtrl imgxPrint 
         Height          =   2040
         Left            =   7920
         TabIndex        =   76
         Top             =   225
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   3598
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind"
         LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl ImgxPrint2 
         Height          =   2040
         Left            =   3420
         TabIndex        =   77
         Top             =   1035
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   3598
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind"
         LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
      End
   End
End
Attribute VB_Name = "sales_ordersfrm"
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
Dim CardTable As ADODB.Recordset
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
aInsert = AddFlag(aInsert, "[DATE]", addDate(xDate.text))
aInsert = AddFlag(aInsert, "[NOTES]", addstring(xNotes.text))
On Error GoTo myerror
con.BeginTrans
If xDoc_No.Tag = DefineMode Then
    xDoc_No.text = IncData("FILE6_51H", "DOC_NO", con, , "00001")
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.text))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(cusername))
    con.Execute addInsert(aInsert, "FILE6_51H")
Else
    aInsert = AddFlag(aInsert, "[USERNAME2]", addstring(cusername))
    con.Execute addUpdate(aInsert, "FILE6_51H", "DOC_NO = " & addstring(xDoc_No.text))
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
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    Dim bNew As Boolean
    bNew = grid1.Row = grid1.Rows - 1
    grid1.TextMatrix(grid1.Row, 1) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    If Not bNew Then
        CellPos 13, grid1.Row, grid1.Col
    End If
    GrdDesc grid1.TextMatrix(grid1.Row, 1), grid1.Row
    Grid1_AfterEdit grid1.Row, grid1.Col
ElseIf ActiveControl.Name = CmdInform.Name Then
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

Private Sub CmdDel_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) <> vbOK Then
    Exit Sub
End If

con.BeginTrans
con.Execute "Delete  From FILE6_51 where Doc_No = " & MyParn(xDoc_No.text)
con.Execute "Delete  From FILE6_51H where Doc_No = " & MyParn(xDoc_No.text)
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
myreplace
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
       xDate.SetFocus
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
ElseIf KeyCode = 116 And CmdSave.Enabled Then
    cmdSave_Click
End If
End Sub
Private Sub Form_Load()
'On Error GoTo myerror
openCon con
nRound = 0

Set xLogo.RowSource = myRs("SELECT * FROM LOGO_CODES", con)
xLogo.ListField = "Desca"
xLogo.BoundColumn = "Code"

Set grid1.DataSource = DATA11
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
    CalcTotals
    Exit Sub
End If

If Row = grid1.Rows - 1 Then
    myAddItem
    bNew = True
End If

CalcTotals

If myreplace(Row) Then
    If xDoc_No.Tag = DefineMode Then
        openCardTable tbMode.tbFind, xDoc_No.text
    ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        grid1.TextMatrix(Row, grid1.Cols - 1) = get_Id(con)
        CalcTotals
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
    ElseIf xIs_collect.Value = 1 Then
        .Editable = flexEDNone
    ElseIf .Col = 1 Or .Col = 10 Or .Col = 11 Then
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
    DOPRINT
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

If Not IsDate(xDate.text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xCodeDesca.Caption = "" Then
    If Not bIgMsg Then MsgBox "·« ÌÊÃœ ﬂÊœ ⁄„Ì·"
    Exit Function
End If
myValid = True
End Function
Private Sub myload()
xDoc_No.text = CardTable!DOC_NO
xDate.text = myFormat_p(CardTable!Date)
xNotes.text = CardTable!NOTES & ""
xCode.text = CardTable!CODE & ""
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
xDoc_No.text = IncData("FILE6_51H", "DOC_NO", con, , "00001")

xCode.text = ""
xCodeDesca.Caption = ""
xDate.text = myFormat_p(Date)


xTotalQuant.Caption = ""
xTotalItem.Caption = ""
'xTotal.Caption = ""
xNotes.text = ""
panel1(1).Caption = GetComputerName
panel1(2).Caption = cusername

grid1.Rows = 1
myAddItem

fixGrd

Handlecontrols DefineMode
CalcTotals

CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub Handlecontrols(nMode)
bEdit = True
bEditRecord = bEdit
cmdNewInv.Enabled = bEdit And nMode = LoadMode

xDoc_No.Tag = nMode


CmdSave.Enabled = bEditRecord
CmdDel.Enabled = (bEditRecord Or bDelRecord) And nMode = LoadMode

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
ElseIf xIs_collect.Value = 1 Then
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
        
    CalcTotals
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
grid1.TextMatrix(Row, 10) = "1"
grid1.TextMatrix(Row, 11) = IIf(optPrice(0).Value, cm.Parameters("@COST2").Value, cm.Parameters("@PRICE").Value)
grid1.TextMatrix(Row, 14) = cm.Parameters("@COST").Value
'grid1.TextMatrix(Row, 15) = fnBalance(cm.Parameters("@ITEM").Value, con, xStore.BoundText) & ""
GrdDesc = True
End With
End Function
Private Function CalcTotals(Optional nMode As Integer = 0)
Dim nTotalQuant As Double, nTotalQuant2 As Double, nTotalItem As Double, nTotalItemNoDiscount As Double, nTotalDiscount As Double, nTotalMin As Double
Dim nDisountRow As Double
With grid1
Dim i As Long
For i = 1 To grid1.Rows - 2
    .TextMatrix(i, 13) = .ValueMatrix(i, 10) * .ValueMatrix(i, 11)
    nTotalQuant = nTotalQuant + grid1.ValueMatrix(i, 10)
    nTotalItem = mRound(nTotalItem + (.ValueMatrix(i, 10) * .ValueMatrix(i, 11)), 6)
Next
xTotalQuant.Caption = Myvalue(nTotalQuant)
xTotalItem.Caption = Myvalue(mRound(nTotalItem))


'bIg = True
'If Val(xTotalItem.Caption) <> 0 Then
'    If Round(Val(xRate.text), nRound) <> Round(Val(xDiscount.text) / Val(xTotalItem.Caption) * 100, nRound) Then
'        xRate.text = Myvalue(Round((Val(xDiscount.text) / Val(xTotalItem.Caption)) * 100, nRound))
'    End If
'Else
'    xRate.text = ""
'End If
'
'
'If IsDate(XDATESALES1.text) Or IsDate(XDATESALES2.text) Then
'    xRate_Tax.text = "14"
'    xtax.text = mRound(Val(xTotalItem.Caption) * 0.14)
'Else
'    If (Val(xTotalItem.Caption) - Val(xDiscount.text)) <> 0 Then
'        If Round(Val(xRate_Tax.text), nRound) <> Round(Val(xtax.text) / (Val(xTotalItem.Caption) - Val(xDiscount.text)) * 100, nRound) Then
'            xRate_Tax.text = Myvalue(Round((Val(xtax.text) / (Val(xTotalItem.Caption) - Val(xDiscount.text))) * 100, nRound))
'        End If
'    Else
'        xRate_Tax.text = ""
'    End If
'End If
'For i = 1 To .Rows - 2
'    If .ValueMatrix(i, 9) > 0 Then
'         nPrice = .ValueMatrix(i, 9) / ((1 + (Val(xRate_Tax.text) / 100)))
'        .TextMatrix(i, 12) = Myvalue(mRound((nPrice - .ValueMatrix(i, 11)) / nPrice, 4) * 100)
'    Else
'        .TextMatrix(i, 12) = ""
'    End If
'Next
'
'xTotal.Caption = mRound((nTotalItem - Val(xDiscount.text)) + Val(xtax.text), 2)
'bIg = False
End With
End Function
Private Sub CardLookup(Optional pFilter As String = "")
Dim Generalarray(5)
Dim listarray(2, 5)
Dim GrdArray(3, 1)
Set Generalarray(0) = Me

Generalarray(1) = "SELECT TOP 1000 " & _
                  " FILE6_51H.DOC_NO ," & _
                  " FORMAT(FILE6_51H.date,'yyyy/M/d')," & _
                  " FILE3_10.DESCA," & _
                  " FILE6_51H.NOTES" & _
                  " FROM FILE6_51H" & _
                  " INNER JOIN FILE3_10 ON FILE6_51H.CODE = FILE3_10.CODE"

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & " WHERE " & cFilter
End If

Generalarray(2) = "Order by DATE DESC,doc_no DESC"
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ «·„” ‰œ"
listarray(0, 1) = "(FILE6_51H.DOC_NO = 'cFilter')"

listarray(1, 0) = "«· «—ÌŒ"
listarray(1, 1) = "(##FILE6_51H.DATE##)"

listarray(2, 0) = "«·⁄„Ì·"
listarray(2, 1) = "(%%FILE3_10.DESCA%%)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1200

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 1300

GrdArray(2, 0) = "«·⁄„Ì·"
GrdArray(2, 1) = 4500

GrdArray(3, 0) = "«·»Ì«‰"
GrdArray(3, 1) = 3500


searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub


Private Sub xDiscount_offer_Change()
lblOffer.Visible = Val(xDiscount_offer.Caption) <> 0
xtotalitem_offer.Visible = Val(xDiscount_offer.Caption) <> 0
End Sub

Private Sub xdoc_replace_Change()
xdoc_replace.Visible = xdoc_replace.Caption <> ""
If xdoc_replace.Caption <> "" Then
    xType.Caption = "«” »œ«·"
Else
    xType.Caption = ""
End If
End Sub
Private Sub xdoc_ret_Change()
cmdReplace.Enabled = IsDgt(xdoc_ret.text) And xDoc_No.Tag = DefineMode And (Len(Trim(xdoc_ret.text)) >= 11)
cmdRefund.Enabled = IsDgt(xdoc_ret.text) And xDoc_No.Tag = DefineMode And (Len(Trim(xdoc_ret.text)) >= 11)
End Sub

Private Sub xdoc_ret_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then cmd_showdoc_Click
End Sub

Private Sub xIsRet_Click()
If bIg Then Exit Sub
If MsgBox("«·”„«Õ »„— Ã⁄ ··›« Ê—…", vbYesNo + vbDefaultButton2) = vbYes Then
    con.Execute " UPDATE FILE6_51H SET ISRET = " & XISRET.Value & " , ISNEW = 1 WHERE DOC_NO = " & MyParn(xDoc_No.text), nRec
    If nRec = 1 Then MsgBox "  „  ⁄œÌ· «·„” ‰œ "
End If
End Sub

Private Sub xMan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (Not xMan.MatchedWithList) And IsDgt(xMan.text) Then
        Dim sCodeEmp As String
        sCodeEmp = xMan.BoundText
        xMan.BoundText = RetZero(xMan.text, 4)
        If Not xMan.MatchedWithList Then
            If Not xMan.MatchedWithList Then
                xMan.BoundText = GetField("SELECT CODE FROM FILE6_25 WHERE ID_EMP = " & MyParn(sCodeEmp), con)
            End If
        End If
    End If
End If
End Sub
Private Sub xMAN_Validate(Cancel As Boolean)
If Not xMan.MatchedWithList Then
    Cancel = True
ElseIf xMan.BoundText = "" Then
    Cancel = True
Else
    grid1.TextMatrix(grid1.Rows - 1, 15) = xMan.BoundText
End If
End Sub

Private Sub XPHONE_Click()
Set Phonesfrm.myform = Me
Set Phonesfrm.con = con
Phonesfrm.Show 1
End Sub

Private Sub xPrinted_Click()
If bIg Then Exit Sub
'If xPrinted.Value = 1 Then Exit Sub
con.Execute "update FILE6_51h set printed = " & xPrinted.Value & ",closed = " & xPrinted.Value & " from FILE6_51h where doc_no = " & MyParn(xDoc_No.text)
MsgBox " „ › Õ «·›« Ê—… »‰Ã«Õ"
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then
        myDefine
    End If
End If
End Sub
Sub OrderNoLook()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO, " & _
                  " FORMAT([DATE],'yyyy/M/d')," & _
                  " NOTES " & _
                  " FROM FILE6_51H WHERE doc_sales IS NULL AND ISPOSTED = 0"

If xCode.text <> "" Then
    Generalarray(1) = Generalarray(1) & " AND FILE6_51H.CODE = " & MyParn(xCode.text)
End If

Generalarray(2) = "Order by Date , DOC_NO "
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ"
listarray(0, 1) = "(@@Doc_No@@6 OR " & _
                  " ##[DATE]##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 1300

GrdArray(2, 0) = "«·»Ì«‰"
GrdArray(2, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchOrder.sCaption = "«” ⁄·«„ «·ÿ·»Ì« "
oSearchOrder.Show 1
End Sub

Private Sub xRate_Change()
If bIg Then Exit Sub
bIg = True
If Val(xTotalItem.Caption) <> 0 Then
    If Round(Val(xRate.text), 0) <> Round(Val(xDiscount.text) / Val(xTotalItem.Caption) * 100, 0) Then
        xDiscount.text = Round((Val(xRate.text) * Val(xTotalItem.Caption)) / 100, 2)
    End If
Else
    xDiscount.text = ""
End If
CalcTotals
bIg = False
End Sub
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Sub fixGrd()
With grid1
'                 0         1          2         3         4          5             6          7          8         9          10        11           12           13          14            15
.FormatString = "„.|" & "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "”⁄—|" & "⁄œœ|" & "”⁄— «·»Ì⁄|" & "‰”»… Œ’„|" & "«·≈Ã„«·Ï|" & "«· ﬂ·›…|"
.RowHeight(0) = 600
.WordWrap = True


.ColWidth(0) = 600
.ColWidth(1) = 1000
.ColWidth(2) = 0
.ColWidth(3) = 2000
.ColWidth(4) = 0
.ColWidth(5) = 2000
.ColWidth(6) = 4000
.ColWidth(7) = 1000
.ColWidth(8) = 1600
.ColWidth(9) = 1000
.ColWidth(10) = 800
.ColWidth(11) = 1000
.ColWidth(12) = 1000
.ColWidth(13) = 1300
.ColWidth(14) = 1300
.ColWidth(15) = 900
'.ColHidden(.Cols - 2) = True
.ColHidden(.Cols - 1) = True
If xIs_collect.Value = 1 Then
    .ColFormat(9) = "0.00"
    .ColFormat(11) = "0.00"
    .ColFormat(12) = "0.00"
    .ColFormat(13) = "0.00"
    .ColFormat(14) = "0.00"
End If
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
'.Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter
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
    con.Execute "Delete  From FILE6_51 where id = " & grid1.TextMatrix(Row, grid1.Cols - 1)
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
        aInsert = AddFlag(aInsert, "PRICE_C", .ValueMatrix(i, 9))
        aInsert = AddFlag(aInsert, "QUANT", .ValueMatrix(i, 10))
        aInsert = AddFlag(aInsert, "PRICE", .ValueMatrix(i, 11))
        aInsert = AddFlag(aInsert, "COST", .ValueMatrix(i, 14))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.text))
            con.Execute addInsert(aInsert, "FILE6_51")
        Else
            con.Execute addUpdate(aInsert, "FILE6_51", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub myLoadGrd()
Dim aPrm As Variant

aPrm = AddFlag(aPrm, "DOC_NO", xDoc_No.text)
Set DATA11.Recordset = cmd("[dbo].[sp_orders]", con, adStoredProc, aPrm).Execute
myAddItem
CalcTotals
fixGrd
End Sub
Private Function mysave() As Boolean

If Not myValid Then Exit Function

CalcTotals

If Not myreplace Then Exit Function

Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
                    
End Function
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = True) As Boolean
With grid1
If Not IsNumeric(.TextMatrix(Row, 1)) Then Exit Function
If .ValueMatrix(Row, 10) = 0 Then Exit Function
End With
validRow = True
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf Col = 1 And grid1.TextMatrix(Row, Col) = "" Then
ElseIf Col < 10 Then
    grid1.Col = 10
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 10)
    grid1.ShowCell grid1.Row, 1
End If
End Sub
Private Sub myAddItem()
grid1.AddItem ""
MakeSerial
End Sub
Private Sub xTotal_Change()
If Not bIg Then
    bIg = True
    xDiscount.text = Round(Val(xTotalItem.Caption) - Val(xTotal.Caption), 2)
    CalcTotals
    bIg = False
End If
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
Private Function retRecords(pDoc_No, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If pDoc_No <> "" Then
    'cString = "SELECT Count(FILE6_51H.DOC_NO) AS records,COUNT(CASE WHEN FILE6_51H.DOC_NO_FLAG <= " & MyParn(pDoc_no) & " THEN 1 END) AS record"
    cString = "SELECT Count(FILE6_51H.DOC_NO) AS records,COUNT(CASE WHEN FILE6_51H.DOC_NO <= " & MyParn(pDoc_No) & " THEN 1 END) AS record"
Else
    cString = "SELECT Count(FILE6_51H.DOC_NO) AS records,0 as record"
End If

cString = cString & " FROM FILE6_51H " & _
                    " INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_51H.STORE"

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
    retFilter = retFilter & " AND FILE6_51H.CODE = " & MyParn(cmdClient.Tag)
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
Private Sub xDate_GotFocus()
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

Dim loctable As New ADODB.Recordset
Set loctable = cmd("SELECT * FROM FILE3_10 WHERE CODE = " & MyParn(xCode.text), con).Execute
If loctable.EOF Then Exit Sub
xCodeDesca.Caption = loctable!DESCA & ""
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_No
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No
End Sub
Private Sub XSTORE_GotFocus()
myGotFocus xStore
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus xStore
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub DOPRINT()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim lPrice As Boolean
Dim nFBal As Double
Dim nPay As Double
Dim nBal As Double

Dim nWidth As Long, nHeight As Long
Dim imageTable As New ADODB.Recordset

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

Dim loctable As New ADODB.Recordset, cString As String
cString = "SELECT FILE6_51H.DOC_NO," & _
          " FILE6_51H.DATE," & _
          " FILE3_10.DESCA AS CODE_DESCA," & _
          " FILE6_51.ITEM," & _
          " FILE6_51.QUANT," & _
          " FILE6_51.PRICE," & _
          " FILE1_10.DESCA AS ITEM_DESCA," & _
          " FILE6_51.TOTAL," & _
          " file1_10.modelfact," & _
          " FACT.DESCA AS FACTDESCA," & _
          " FILE1_10.COLOR," & _
          " FILE1_10.C_COLOR," & _
          " FILE1_10.SCAL," & _
          " FILE1_10.MODEL" & _
          " FROM FILE6_51 " & _
          " INNER JOIN FILE6_51H ON FILE6_51.DOC_NO = FILE6_51H.DOC_NO" & _
          " INNER JOIN FILE1_10 ON FILE6_51.ITEM = FILE1_10.ITEM" & _
          " INNER JOIN FILE3_10 ON FILE6_51H.CODE = FILE3_10.CODE" & _
          " LEFT JOIN FACT ON FACT.CODE = FILE1_10.[FACT]" & _
          " WHERE  FILE6_51.DOC_NO = " & MyParn(xDoc_No.text)
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
             " ORDER BY C_COLOR DESC"
     
    If chkImage.Value = 1 Then
        Set imageTable = cmd(cString, con_image).Execute
        ImgxPrint2.Images.Clear
        If Not imageTable.EOF Then
            If Not IsNull(imageTable!Image) Then
                Set picPrint.Picture = LoadPicture("")
                Set picPrint.DataSource = imageTable
                picPrint.DataField = "image"
                
                'GetNewSize picPrint.Width, picPrint.Height, nWidth, 200
                
                aRet = retDimImg(picPrint.Width, picPrint.Height, 300, 300)
    
                ImgxPrint2.Import.FromMemoryFile (imageTable!Image)
                ImgxPrint2.Effects.Resize retFlag(aRet, "width"), retFlag(aRet, "height")
                'ImgxPrint2.Effects.Resize nWidth, 200
            
                SaveImageToRs ImgxPrint2.Picture, temptable, "picture2"
            End If
        End If
    End If
    
    If imgxPrint.Images.Count > 0 Then
        SaveImageToRs imgxPrint.Picture, temptable, "picture1"
    End If
    
    temptable!str6 = loctable!DOC_NO
    temptable!Date1 = myFormat_p(loctable!Date)
    temptable!STR7 = loctable!CODE_desca
    
    temptable!str2 = loctable!FACTDESCA
    temptable!str1 = loctable!Item
    temptable!str13 = TurnValue(DelZero(loctable!MODELFACT))
    temptable!str3 = loctable!ITEM_DESCA
    temptable!str4 = loctable!color
    temptable!str5 = loctable!SCAL
    temptable!str16 = loctable!MODEL
    temptable!str10 = "⁄—÷ ”⁄—"
    
    temptable!val3 = loctable!Quant
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
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim lPrice As Boolean
Dim nFBal As Double
Dim nPay As Double
Dim nBal As Double

Dim nWidth As Long, nHeight As Long
Dim imageTable As New ADODB.Recordset

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

Dim loctable As New ADODB.Recordset
Dim cString As String
cString = "SELECT FILE1_10.MODEL," & _
            "FILE1_10.MODELFACT," & _
            "FILE1_10.DESCA AS ITEM_DESCA," & _
            "MAX(FILE6_51.PRICE) AS PRICE," & _
            "FILE6_51H.DOC_NO," & _
            "FILE6_51H.[DATE]," & _
            "FACT.DESCA AS FACT_DESCA," & _
            "FILE3_10.DESCA AS CODE_desca," & _
            " (SELECT  STRING_AGG(COLOR, CHAR(13)) AS COLOR FROM (SELECT DISTINCT  COLOR FROM FILE1_10 F where model = FILE1_10.MODEL) FC) AS COLOR," & _
            " (SELECT  STRING_AGG(SCAL, CHAR(13)) AS SCAL FROM (SELECT DISTINCT  SCAL FROM FILE1_10 F where model = FILE1_10.MODEL) FS) AS SCAL" & _
            " From FILE6_51 " & _
            " INNER JOIN FILE6_51H ON FILE6_51.DOC_NO = FILE6_51H.DOC_NO" & _
            " INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE6_51.ITEM" & _
            " INNER JOIN FILE3_10 ON FILE6_51H.CODE = FILE3_10.CODE" & _
            " LEFT JOIN FACT ON FACT.CODE = FILE1_10.[FACT]" & _
            " WHERE FILE6_51H.DOC_NO = " & MyParn(xDoc_No.text)
 
 cString = cString & _
            " GROUP BY FILE6_51H.DOC_NO," & _
             "FILE1_10.MODEL," & _
             "FILE1_10.MODELFACT," & _
             "FILE1_10.DESCA," & _
             "FILE6_51.DOC_NO," & _
             "FILE6_51H.[DATE]," & _
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
             " ORDER BY C_COLOR"
     
        Set imageTable = cmd(cString, con_image).Execute
        ImgxPrint2.Images.Clear
        If Not imageTable.EOF Then
            If Not IsNull(imageTable!Image) Then
                Set picPrint.Picture = LoadPicture("")
                Set picPrint.DataSource = imageTable
                picPrint.DataField = "image"
                
                'GetNewSize picPrint.Width, picPrint.Height, nWidth, 200
                
                aRet = retDimImg(picPrint.Width, picPrint.Height, 300, 300)
    
                ImgxPrint2.Import.FromMemoryFile (imageTable!Image)
                ImgxPrint2.Effects.Resize retFlag(aRet, "width"), retFlag(aRet, "height")
                'ImgxPrint2.Effects.Resize nWidth, 200
            
                SaveImageToRs ImgxPrint2.Picture, temptable, "picture2"
            End If
        End If
    
    If imgxPrint.Images.Count > 0 Then
        SaveImageToRs imgxPrint.Picture, temptable, "picture1"
    End If
    
    temptable!str6 = loctable!DOC_NO
    temptable!Date1 = myFormat_p(loctable!Date)
    temptable!STR7 = loctable!CODE_desca
    temptable!VAL1 = loctable!price
    
    temptable!str2 = loctable!FACT_DESCA
    temptable!MEMO1 = loctable!color
    temptable!MEMO2 = loctable!SCAL
    temptable!str16 = loctable!MODEL
    temptable!str13 = TurnValue(DelZero(loctable!MODELFACT))
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
Private Function openCardTable(Optional pMode As Integer = tbMode.tblast, Optional pDoc_No As String = "", Optional pDoc_no_Filter As String = "") As Boolean
Dim cString As String
Dim cWhere As String
Dim cOrder As String
Dim nRecords As Long, nRecord As Long

If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pDoc_No) = "" And Trim(pDoc_no_Filter) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass


On Error GoTo myerror
cFilter = retFilter

Set CardTable = New ADODB.Recordset
cString = "SELECT TOP 1 FILE6_51H.*," & _
          " FILE3_10.DESCA AS CLIENTDESCA " & _
          " FROM FILE6_51H " & _
          " INNER JOIN FILE3_10 ON FILE3_10.CODE = FILE6_51H.code"
  
If pMode = tbMode.tbFirst Then
    'cOrder = "Order by FILE6_51H.DOC_NO_FLAG"
    cOrder = "Order by FILE6_51H.DOC_NO"
ElseIf pMode = tbMode.tblast Then
    'cOrder = "Order by FILE6_51H.DOC_NO_FLAG DESC"
    cOrder = "Order by FILE6_51H.DOC_NO DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "FILE6_51H.Doc_no = " & MyParn(pDoc_No)
ElseIf pMode = tbMode.tbPrevious Then
    'cWhere = "FILE6_51H.DOC_NO_FLAG < " & MyParn(pDoc_no)
    'cOrder = "Order by FILE6_51H.DOC_NO_FLAG DESC"
    
    cWhere = "FILE6_51H.DOC_NO < " & MyParn(pDoc_No)
    cOrder = "Order by FILE6_51H.DOC_NO DESC"
ElseIf pMode = tbMode.tbNext Then
    'cWhere = "FILE6_51H.DOC_NO_FLAG > " & MyParn(pDoc_no)
    cWhere = "FILE6_51H.DOC_NO > " & MyParn(pDoc_No)
    cOrder = "Order by FILE6_51H.DOC_NO"
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


