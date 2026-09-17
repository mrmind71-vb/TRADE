VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form sales_wholefrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11085
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
   ScaleHeight     =   11085
   ScaleWidth      =   20355
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
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
      Left            =   -8460
      TabIndex        =   26
      Top             =   10890
      Visible         =   0   'False
      Width           =   9285
      Begin VB.Frame fmCollectTrans 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   0
         TabIndex        =   138
         Top             =   0
         Width           =   3660
         Begin VB.TextBox XDATESALES1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1665
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   225
            Width           =   1230
         End
         Begin VB.TextBox XDATESALES2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   360
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   225
            Width           =   1275
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "«· «—ÌŒ"
            Height          =   270
            Index           =   1
            Left            =   2970
            RightToLeft     =   -1  'True
            TabIndex        =   141
            Top             =   270
            Width           =   510
         End
      End
      Begin VB.CheckBox xIs_collect 
         Caption         =   "Check1"
         Height          =   270
         Left            =   360
         TabIndex        =   125
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
         TabIndex        =   43
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   28
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
         TabIndex        =   126
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
         TabIndex        =   120
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
         TabIndex        =   42
         Top             =   1035
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label XBRANCH 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„” Œœ„"
         Height          =   285
         Left            =   0
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   27
         Top             =   405
         Visible         =   0   'False
         Width           =   675
      End
   End
   Begin VB.Frame frPrinter 
      Height          =   4335
      Left            =   -3960
      RightToLeft     =   -1  'True
      TabIndex        =   134
      Top             =   11025
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
         TabIndex        =   135
         Top             =   360
         Width           =   1125
      End
      Begin ImgXCtrl6.ImgXCtrl imgxPrint 
         Height          =   2040
         Left            =   7920
         TabIndex        =   136
         Top             =   225
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   3598
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind"
         LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl ImgxPrint2 
         Height          =   945
         Left            =   3420
         TabIndex        =   137
         Top             =   1035
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1667
         AutoSize        =   -1  'True
         LicenseUserName =   "mrmind"
         LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
      End
      Begin Threed.SSCommand cmdPurchase 
         Height          =   510
         Left            =   0
         TabIndex        =   145
         Top             =   0
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
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
         Caption         =   " ÕÊÌ· ·„‘ —Ì«   ÊﬂÌ·« "
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   3510
      TabIndex        =   131
      Top             =   -45
      Width           =   1815
      Begin VB.CheckBox chkNoBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " Ã«Â· «·—’Ìœ"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   270
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   10080
      TabIndex        =   127
      Top             =   -90
      Width           =   2400
      Begin Threed.SSCommand CMD_ADD_RECEVD 
         Height          =   510
         Left            =   45
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   180
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
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
         Caption         =   "≈÷«›… „‰ «” ·«„ ÿ·»Ì…"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
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
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   1575
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox xDoc_No 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   -2430
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   630
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Frame fmE_Invoice 
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
      Left            =   90
      TabIndex        =   112
      Top             =   900
      Width           =   8700
      Begin Threed.SSCommand cmdSign 
         Height          =   330
         Left            =   1890
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   180
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         _Version        =   196610
         ForeColor       =   0
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
         Caption         =   "«—”«· ›« Ê—…  ›’Ì·Ì…"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdSignGroup 
         Height          =   330
         Left            =   90
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   180
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   582
         _Version        =   196610
         ForeColor       =   0
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
         Caption         =   "«—”«· ›« Ê—… «Ã„«·Ì…"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "›« Ê—… «·Ìﬂ —Ê‰Ì…"
         Height          =   285
         Left            =   7290
         RightToLeft     =   -1  'True
         TabIndex        =   115
         Top             =   225
         Width           =   1290
      End
      Begin VB.Label xinv_Tax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   114
         Top             =   180
         Width           =   3495
      End
   End
   Begin VB.Frame fmTransTax 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   5355
      TabIndex        =   102
      Top             =   -90
      Width           =   4695
      Begin Threed.SSCommand cmdTaxInv 
         Height          =   330
         Left            =   135
         TabIndex        =   104
         Top             =   225
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
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
         Caption         =   " ÕÊÌ· «·Ì"
         TagVariant      =   "«Œ «— «·„Œ“‰"
         ButtonStyle     =   3
      End
      Begin VB.Label xdoc_no_tax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1755
         TabIndex        =   111
         Top             =   225
         Width           =   1545
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "›« Ê—… ÷—Ì»Ì…"
         Height          =   240
         Index           =   2
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   103
         Top             =   270
         Width           =   1170
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   0
      ScaleHeight     =   1725
      ScaleWidth      =   20355
      TabIndex        =   62
      Top             =   8700
      Width           =   20355
      Begin VB.Frame fmCollect 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   135
         TabIndex        =   123
         Top             =   -45
         Width           =   1725
         Begin VB.CheckBox chkTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "›« Ê—… „Ã„⁄…"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   124
            TabStop         =   0   'False
            Top             =   270
            Width           =   1455
         End
      End
      Begin VB.Frame fmShowTotal 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   1890
         TabIndex        =   118
         Top             =   -45
         Visible         =   0   'False
         Width           =   1725
         Begin VB.CheckBox chkTotalInv 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "›« Ê—… «Ã„«·Ì…"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   270
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   3645
         TabIndex        =   98
         Top             =   -45
         Width           =   11715
         Begin VB.OptionButton optType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«–‰"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   9930
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   270
            Width           =   645
         End
         Begin VB.OptionButton optType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "›Ê« Ì— ÷—Ì»Ì… „— Ã⁄« "
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   109
            Top             =   270
            Width           =   2175
         End
         Begin VB.OptionButton optType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "›Ê« Ì— ÷—Ì»Ì… „»Ì⁄« "
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   5
            Left            =   2535
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   270
            Width           =   1995
         End
         Begin VB.OptionButton optType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "›Ê« Ì— ÷—Ì»Ì…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   4845
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   270
            Width           =   1590
         End
         Begin VB.OptionButton optType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«–‰ „— Ã⁄« "
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   6750
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   270
            Width           =   1320
         End
         Begin VB.OptionButton optType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«–‰ „»Ì⁄« "
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   8385
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   270
            Width           =   1230
         End
         Begin VB.OptionButton optType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«·ﬂ·"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   10890
            TabIndex        =   99
            Top             =   270
            Value           =   -1  'True
            Width           =   690
         End
      End
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   1080
         Width           =   3525
         Begin Threed.SSCommand cmdFirst 
            Height          =   420
            Left            =   2610
            TabIndex        =   74
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
            Picture         =   "sales_whole.frx":0000
            Caption         =   "√Ê·"
            ButtonStyle     =   3
            PictureAlignment=   10
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "sales_whole.frx":21A7
         End
         Begin Threed.SSCommand cmdPrevious 
            Height          =   420
            Left            =   1710
            TabIndex        =   75
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
            Picture         =   "sales_whole.frx":41EE
            Caption         =   "”«»ﬁ"
            ButtonStyle     =   3
            PictureAlignment=   10
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "sales_whole.frx":62D9
         End
         Begin Threed.SSCommand cmdNext 
            Height          =   420
            Left            =   855
            TabIndex        =   76
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
            Picture         =   "sales_whole.frx":82D3
            Caption         =   "·«Õﬁ"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "sales_whole.frx":A3E4
         End
         Begin Threed.SSCommand cmdLast 
            Height          =   420
            Left            =   45
            TabIndex        =   77
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
            Picture         =   "sales_whole.frx":C3DE
            Caption         =   "√ŒÌ—"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "sales_whole.frx":E602
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   585
         Width           =   3255
         Begin VB.TextBox txtModel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   90
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   92
            TabStop         =   0   'False
            Tag             =   "1"
            Top             =   270
            Width           =   2040
         End
         Begin VB.TextBox txtBarcode 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   90
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   91
            TabStop         =   0   'False
            Tag             =   "1"
            Top             =   675
            Width           =   2040
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "—ﬁ„ „ÊœÌ·"
            Height          =   285
            Left            =   2205
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   315
            Width           =   825
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "»«—ﬂÊœ"
            Height          =   285
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   720
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
         Height          =   1680
         Left            =   15390
         TabIndex        =   78
         Top             =   0
         Width           =   4920
         Begin VB.TextBox xRate_Tax 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00400000&
            Height          =   330
            Left            =   1575
            MaxLength       =   10
            TabIndex        =   96
            Top             =   900
            Width           =   555
         End
         Begin VB.TextBox xtax 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00400000&
            Height          =   330
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   95
            Top             =   900
            Width           =   1410
         End
         Begin VB.TextBox xRate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00400000&
            Height          =   330
            Left            =   1575
            MaxLength       =   10
            TabIndex        =   80
            Top             =   540
            Width           =   555
         End
         Begin VB.TextBox xDiscount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00400000&
            Height          =   330
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   79
            Top             =   540
            Width           =   1410
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "«·÷—Ì»…"
            Height          =   270
            Left            =   3645
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   945
            Width           =   600
         End
         Begin VB.Shape Shape1 
            Height          =   330
            Left            =   1575
            Top             =   1260
            Width           =   1995
         End
         Begin VB.Label xTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1575
            TabIndex        =   87
            Top             =   1260
            Width           =   1995
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "’«›Ì «·›« Ê—…"
            Height          =   270
            Left            =   3645
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   1305
            Width           =   1095
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Œ’„ ›« Ê—…"
            Height          =   270
            Left            =   3645
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   585
            Width           =   900
         End
         Begin VB.Label xTotalItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2115
            TabIndex        =   84
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„… «’‰«›"
            Height          =   270
            Left            =   3645
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   225
            Width           =   900
         End
         Begin VB.Label xTotalQuant 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   135
            TabIndex        =   82
            Top             =   180
            Width           =   825
         End
         Begin VB.Label lblQuant 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ ﬁÿ⁄"
            Height          =   270
            Left            =   1035
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   225
            Width           =   630
         End
      End
      Begin VB.Frame fmMonth 
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
         Height          =   645
         Left            =   9945
         TabIndex        =   67
         Top             =   585
         Width           =   5415
         Begin VB.CheckBox chkDay 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "›Ê« Ì— «·ÌÊ„"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   270
            Width           =   1230
         End
         Begin MSDataListLib.DataCombo xYear 
            Height          =   330
            Left            =   3465
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   225
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Style           =   2
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
         Begin MSDataListLib.DataCombo xMonth 
            Height          =   330
            Left            =   1530
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   225
            Width           =   870
            _ExtentX        =   1535
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
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·‘Â—"
            Height          =   330
            Left            =   2385
            TabIndex        =   72
            Tag             =   "Color"
            Top             =   270
            Width           =   555
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·”‰…"
            Height          =   285
            Left            =   4725
            TabIndex        =   71
            Tag             =   "Color"
            Top             =   270
            Width           =   510
         End
      End
      Begin VB.Frame fmDate 
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
         Height          =   645
         Left            =   6975
         TabIndex        =   63
         Top             =   585
         Width           =   2940
         Begin VB.TextBox xdate_filter 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   720
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   180
            Width           =   1455
         End
         Begin Threed.SSCommand DAYALL1 
            Height          =   375
            Left            =   90
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   180
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   661
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
            Picture         =   "sales_whole.frx":106D3
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            PictureDisabledFrames=   1
            ShapeSize       =   1
            PictureDisabled =   "sales_whole.frx":128D5
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   0
            Left            =   2340
            TabIndex        =   66
            Tag             =   "Color"
            Top             =   225
            Width           =   555
         End
      End
      Begin Threed.SSCommand cmdStore 
         Height          =   375
         Left            =   11070
         TabIndex        =   88
         Top             =   1305
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   661
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
         Caption         =   "«Œ «— «·„Œ“‰"
         TagVariant      =   "«Œ «— «·„Œ“‰"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClient 
         Height          =   375
         Left            =   6975
         TabIndex        =   89
         Top             =   1305
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   661
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
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         Height          =   600
         Left            =   135
         TabIndex        =   129
         Top             =   495
         Width           =   3435
         Begin VB.CheckBox chkNoTax 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "›Ê« Ì— ÷—Ì»Ì… €Ì— „—Õ·…"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   945
            RightToLeft     =   -1  'True
            TabIndex        =   130
            TabStop         =   0   'False
            Top             =   225
            Width           =   2400
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   5310
      TabIndex        =   58
      Top             =   1485
      Width           =   3480
      Begin Threed.SSCommand cmdCsv 
         Height          =   510
         Left            =   1890
         TabIndex        =   59
         Top             =   180
         Width           =   1545
         _ExtentX        =   2725
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
         Picture         =   "sales_whole.frx":149F1
         Caption         =   "”Õ» „‰ „·›"
         ButtonStyle     =   4
         PictureAlignment=   9
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   510
         Left            =   45
         TabIndex        =   60
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
         Picture         =   "sales_whole.frx":16E9F
         Caption         =   " ÕÊÌ· «·Ì „·› "
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame fmPrice 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2205
      TabIndex        =   55
      Top             =   1485
      Width           =   3075
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
         TabIndex        =   57
         Top             =   270
         Width           =   1275
      End
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
         TabIndex        =   56
         Top             =   270
         Value           =   -1  'True
         Width           =   1050
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
      TabIndex        =   20
      Top             =   11070
      Visible         =   0   'False
      Width           =   10725
      Begin VB.CommandButton cmd_showdoc 
         Caption         =   "⁄—÷  ›’Ì·Ï «·»Ê‰"
         Height          =   450
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   21
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
      Height          =   1275
      Index           =   4
      Left            =   8820
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   945
      Width           =   1230
      Begin Threed.SSCommand cmdSave 
         Height          =   555
         Left            =   45
         TabIndex        =   11
         Top             =   135
         Width           =   1140
         _ExtentX        =   2011
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
         Picture         =   "sales_whole.frx":19309
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_whole.frx":1BC2E
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   510
         Left            =   45
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   720
         Width           =   1140
         _ExtentX        =   2011
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
         Picture         =   "sales_whole.frx":1E482
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_whole.frx":205E2
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   14625
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   -90
      Width           =   5685
      Begin Threed.SSCommand cmdInform 
         Height          =   465
         Left            =   4545
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   820
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
         Picture         =   "sales_whole.frx":228CF
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_whole.frx":24C9A
      End
      Begin Threed.SSCommand cmdNewInv 
         Height          =   465
         Left            =   3420
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   820
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
         Picture         =   "sales_whole.frx":26D43
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_whole.frx":28D4B
      End
      Begin Threed.SSCommand cmddel 
         Height          =   465
         Left            =   2295
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   820
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
         Picture         =   "sales_whole.frx":2AD02
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_whole.frx":2D49E
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   465
         Left            =   45
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   180
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   820
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
         Picture         =   "sales_whole.frx":2F932
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   465
         Left            =   1125
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   180
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   820
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
         Picture         =   "sales_whole.frx":31C55
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "sales_whole.frx":33FCB
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
      Left            =   900
      Top             =   1755
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
      TabIndex        =   49
      Top             =   10620
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
         TabIndex        =   50
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
         TabIndex        =   51
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
         TabIndex        =   52
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
         TabIndex        =   53
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
         TabIndex        =   54
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   61
      Top             =   10425
      Visible         =   0   'False
      Width           =   20355
      _ExtentX        =   35904
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VSFlex7Ctl.VSFlexGrid grdError 
      Height          =   6135
      Left            =   45
      TabIndex        =   10
      Top             =   2295
      Visible         =   0   'False
      Width           =   20265
      _cx             =   35745
      _cy             =   10821
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
      BackColorFixed  =   16576
      ForeColorFixed  =   16777215
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
      Cols            =   3
      FixedRows       =   1
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
      Height          =   1680
      Left            =   10080
      TabIndex        =   15
      Top             =   585
      Width           =   10230
      Begin VB.TextBox xDate_due 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
      End
      Begin VB.TextBox xpo_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3510
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   2715
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   7920
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   540
         Width           =   1230
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
         TabIndex        =   14
         Top             =   2460
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3510
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   900
         Width           =   2715
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1395
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   1230
      End
      Begin MSDataListLib.DataCombo xStore 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   495
         TabIndex        =   5
         Top             =   540
         Width           =   2130
         _ExtentX        =   3757
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
      Begin Threed.SSCommand cmdStoreLookup 
         Height          =   330
         Left            =   90
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   540
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
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
         Caption         =   "..."
         ButtonStyle     =   2
         BevelWidth      =   0
      End
      Begin MSDataListLib.DataCombo xLogo 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   7290
         TabIndex        =   7
         Top             =   900
         Width           =   1860
         _ExtentX        =   3281
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
      Begin MSDataListLib.DataCombo xman 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   90
         TabIndex        =   9
         Top             =   900
         Width           =   2535
         _ExtentX        =   4471
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
      Begin Threed.SSCommand xPrice_list 
         Height          =   330
         Left            =   3510
         TabIndex        =   147
         Top             =   1260
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   582
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
         ButtonStyle     =   2
         BevelWidth      =   0
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·‰Ê⁄"
         Height          =   270
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   149
         Top             =   1260
         Width           =   405
      End
      Begin VB.Label lblInv_type 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   90
         TabIndex        =   148
         Top             =   1260
         Width           =   2535
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁ«∆„… «·”⁄—"
         Height          =   270
         Left            =   9225
         RightToLeft     =   -1  'True
         TabIndex        =   146
         Top             =   1305
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„‰œÊ»"
         Height          =   270
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   144
         Top             =   900
         Width           =   600
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·‘⁄«—"
         Height          =   270
         Left            =   9225
         RightToLeft     =   -1  'True
         TabIndex        =   133
         Top             =   945
         Width           =   480
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ ÿ·»Ì… "
         Height          =   240
         Index           =   1
         Left            =   6255
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   225
         Width           =   855
      End
      Begin VB.Label xDoc_no2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7920
         TabIndex        =   0
         Tag             =   "2"
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3510
         TabIndex        =   48
         Top             =   540
         Width           =   4380
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "«·⁄„Ì·"
         Height          =   270
         Left            =   9315
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   585
         Width           =   450
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Œ“‰"
         Height          =   270
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   585
         Width           =   540
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ „” ‰œ"
         Height          =   285
         Left            =   9270
         RightToLeft     =   -1  'True
         TabIndex        =   16
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
         TabIndex        =   19
         Top             =   2430
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "„·«ÕŸ« "
         Height          =   270
         Left            =   6345
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   990
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
         Height          =   270
         Index           =   0
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   180
         Width           =   510
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   5640
      Left            =   90
      TabIndex        =   117
      Top             =   2295
      Visible         =   0   'False
      Width           =   20220
      _cx             =   35666
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
      BackColorFixed  =   12648384
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6180
      Left            =   90
      TabIndex        =   13
      Top             =   2295
      Width           =   20220
      _cx             =   35666
      _cy             =   10901
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
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   12555
      TabIndex        =   142
      Top             =   -90
      Width           =   2040
      Begin VB.CheckBox chkImage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ÿ»«⁄… »’Ê—"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   540
         TabIndex        =   143
         Top             =   270
         Value           =   1  'Checked
         Width           =   1320
      End
   End
   Begin VB.Label xinv_type 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   -270
      RightToLeft     =   -1  'True
      TabIndex        =   105
      Top             =   1035
      Visible         =   0   'False
      Width           =   2355
   End
End
Attribute VB_Name = "sales_wholefrm"
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
Dim rsInv_Type As ADODB.RecordSet
Dim CardTable As ADODB.RecordSet
Dim oSearchItem As New Search_abd, oSearchDoc  As New Search_abd
Dim osearchClient As New Search_abd
Dim oSearchStoreLook As New Search_abd
Dim oSearchClientLook As New Search_abd
Dim oSearchList As New Search_abd
Dim oSearchStore As New Search_abd
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
aInsert = AddFlag(aInsert, "[DATE_DUE]", addDate(xDate_due.text))
aInsert = AddFlag(aInsert, "STORE", addstring(xStore.BoundText))
aInsert = AddFlag(aInsert, "MAN", addstring(xMan.BoundText))
aInsert = AddFlag(aInsert, "LOGO", addvalue(xLogo.BoundText))
aInsert = AddFlag(aInsert, "[PO_NO]", addstring(xpo_no.text))
aInsert = AddFlag(aInsert, "[NOTES]", addstring(xNotes.text))
aInsert = AddFlag(aInsert, "[RATE]", Val(xRate.text))
aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(xDiscount.text))
aInsert = AddFlag(aInsert, "[ISINVOICE]", XISINVOICE.Value)
aInsert = AddFlag(aInsert, "[PRICE_LIST]", addvalue(xPrice_list.Tag))
'aInsert = AddFlag(aInsert, "[DATESALES1]", addstring(XDATESALES1.text))
'aInsert = AddFlag(aInsert, "[DATESALES2]", addstring(XDATESALES2.text))
aInsert = AddFlag(aInsert, "[BRANCH]", addstring("00"))
On Error GoTo myerror
con.BeginTrans
If xDoc_No.Tag = DefineMode Then
    xDoc_No.text = addNewDoc
    
    xDoc_no2.Caption = Mid(xDoc_No.text, 7, 6)
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.text))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(cusername))
    aInsert = AddFlag(aInsert, "[USER_IP]", addstring(GetComputerName))
    con.Execute addInsert(aInsert, "FILE6_20H")
Else
    con.Execute addUpdate(aInsert, "FILE6_20H", "DOC_NO = " & addstring(xDoc_No.text))
End If
If Row > 0 Then myreplaceGrd Row
con.CommitTrans
myreplace = True
Exit Function
myerror:
prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    If Not xStore.MatchedWithList Then
        MsgBox "«·„Œ“‰ €Ì— „”Ã·"
        Exit Sub
    End If
    Dim bNew As Boolean
    bNew = grid1.Row = grid1.Rows - 1
    grid1.TextMatrix(grid1.Row, 1) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    If Not bNew Then
        CellPos 13, grid1.Row, grid1.Col
    End If
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
ElseIf ActiveControl.Name = cmdStoreLookup.Name Then
    xStore.BoundText = oSearchStoreLook.grid1.TextMatrix(oSearchStoreLook.grid1.Row, 0)
    oSearchStoreLook.Hide
    xStore.SetFocus
ElseIf ActiveControl.Name = cmdStore.Name Then
    If oSearchStore.grid1.TextMatrix(oSearchStore.grid1.Row, 0) = "" Then
        cmdStore.Tag = ""
        cmdStore.Caption = cmdStore.TagVariant
    Else
        cmdStore.Tag = oSearchStore.grid1.TextMatrix(oSearchStore.grid1.Row, 0)
        cmdStore.Caption = oSearchStore.grid1.TextMatrix(oSearchStore.grid1.Row, 1)
    End If
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then myDefine
    End If
    oSearchStore.Hide
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
ElseIf ActiveControl.Name = CMD_ADD_RECEVD.Name Then
    orders_soldfrm.sDoc_no_order = oSearchOrder.grid1.TextMatrix(oSearchOrder.grid1.Row, 0)
    orders_soldfrm.sinv_type = xinv_type.Caption
    Unload oSearchOrder
    
    Set orders_soldfrm.myForm = Me
    orders_soldfrm.Show 1
ElseIf ActiveControl.Name = xPrice_list.Name Then
    If oSearchList.grid1.TextMatrix(oSearchList.grid1.Row, 0) = "" Then
        xPrice_list.Tag = ""
        xPrice_list.Caption = xPrice_list.TagVariant
        con.Execute "UPDATE FILE6_20H " & _
                    "SET FILE6_20H.PRICE_LIST = NULL " & _
                    "WHERE FILE6_20H.DOC_NO =" & MyParn(xDoc_No.text)
        Unload oSearchList
    Else
         xPrice_list.Tag = oSearchList.grid1.TextMatrix(oSearchList.grid1.Row, 0)
         xPrice_list.Caption = oSearchList.grid1.TextMatrix(oSearchList.grid1.Row, 1)
         Unload oSearchList
         If MsgBox(" €ÌÌ— «”⁄«— «·›« Ê—… «·Ì ”⁄— «·ﬁ«∆„…", vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
         If ChangePriceList Then
            Inform " „  €ÌÌ— ”⁄— «·›« Ê—… «·Ì «”⁄«— " & oSearchList.grid1.TextMatrix(oSearchList.grid1.Row, 1)
            myLoadGrd
         End If
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub cmdBranch_Click()
End Sub

Private Sub chkDay_Click()
myUndo
End Sub

Private Sub chkNoTax_Click()
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then
        myDefine
    End If
End If
End Sub

Private Sub chkTotal_Click()
If bIg Then Exit Sub
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then
        myDefine
    End If
End If
End Sub

Private Sub chkTotalInv_Click()
If bIg Then Exit Sub
If chkTotalInv.Value = 1 Then
    grid2.Visible = True
    grid1.Visible = False
    myloadgrd2
Else
    grid1.Visible = True
    grid2.Visible = False
    myLoadGrd
End If
End Sub

Private Sub CMD_ADD_RECEVD_Click()
If Not myValid Then Exit Sub
'ClientOrderLookup Me, oSearchOrder, "isPosted = 0"
End Sub

Private Sub cmdClient_Click()
CLIENTLOOKUP Me, osearchClient, , , IIf(cmdClient.Tag = "", "", "ﬂ· «·⁄„·«¡")
End Sub

Private Sub cmdCSV_Click()
If Not myValid Then Exit Sub
If grid1.Rows > 2 Then
    MsgBox " ÊÃœ ”Ã·«  ›Ï «·„” ‰œ"
    Exit Sub
End If
If getCSV Then
    myLoadGrd
End If
End Sub

Private Sub cmdDate_Click()
myUndo
End Sub

Private Sub CmdDel_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) <> vbOK Then
    Exit Sub
End If

con.BeginTrans
con.Execute "Delete  From FILE6_20 where Doc_No = " & MyParn(xDoc_No.text)
con.Execute "UPDATE FILE6_20H SET" & _
             " FILE6_20H.DOC_NO_TAX = NULL " & _
             " FROM FILE6_20H WHERE FILE6_20H.DOC_NO_TAX = " & MyParn(xDoc_No.text)
con.Execute "Delete  From FILE6_20H where Doc_No = " & MyParn(xDoc_No.text)
con.CommitTrans
    
'If Not openCardTable(tbMode.tbPrevious, xdoc_no_flag.Caption) Then
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
    If grdError.Visible Then
        ToFileExelNew grdError, , , aRow, , 1, , , , , , Me, Array("√Œÿ«¡ ”Õ» »Ì«‰«  ›« Ê—… „»Ì⁄«  —ﬁ„ : " & xDoc_no2.Caption, "«·⁄„Ì· : " & xCodeDesca.Caption, "«· «—ÌŒ : " & xDate.text)
    Else
        ToFileExelNew grid1, , , aRow, , 0.9, , , , , , Me, Array("›« Ê—… „»Ì⁄«  —ﬁ„ : " & xDoc_no2.Caption, "«·⁄„Ì· : " & xCodeDesca.Caption, "«· «—ÌŒ : " & xDate.text)
    End If

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

Private Sub cmdPurchase_Click()
    Dim cDoc As String, cStore As String, cCode As String
    cCode = "000"
    cStore = GetDesca("SELECT BRANCH FROM FILE4_10 WHERE CUST = " & MyParn(xCode.text), con)
    If cStore = "" Then
        MsgBox " «·⁄„Ì· ·Ì”  ÊﬂÌ· "
        Exit Sub
    End If
    Dim aInsert As Variant
    If MsgBox("≈·Ï „” ‰œ „‘ —Ì«   ÊﬂÌ·«  ", vbYesNo + vbDefaultButton2) = vbYes Then
        aInsert = AddFlag(Empty, "[DATE]", DateSq(Date))
        aInsert = AddFlag(aInsert, "[mosm]", addstring(cPMosm))
        aInsert = AddFlag(aInsert, "[STORE]", addstring(cStore))
        aInsert = AddFlag(aInsert, "[code]", addstring(cCode))
        aInsert = AddFlag(aInsert, "[isdate]", "getdate()")
        aInsert = AddFlag(aInsert, "[USERSEND]", addstring(cusername))
        aInsert = AddFlag(aInsert, "[Notes_Doc]", addstring(" „»Ì⁄«  Junior "))
        aInsert = AddFlag(aInsert, "[TRANS_DOC]", addstring(xDoc_No.text))
        aInsert = AddFlag(aInsert, "[ISNEW1]", 1)
        cDoc = Newflag_PurchBr("FR7_20H", "DOC_NO", cStore, con)
        aInsert = AddFlag(aInsert, "DOC_NO", addstring(cDoc))
        con.Execute addInsert(aInsert, "FR7_20h")
        
        con.Execute "  INSERT INTO FR7_20 (DOC_NO , quant, ITEM, price )  SELECT  " & addstring(cDoc) & " , quant, FILE6_20.ITEM, file6_20.price From FILE6_20 inner join file1_10 on FILE6_20.item = file1_10.item WHERE DOC_NO = " & MyParn(xDoc_No.text)
        Inform " „  ÕÊÌ· ≈·Ï „” ‰œ „‘ —Ì«   ÊﬂÌ·«  —ﬁ„ " & cDoc
    End If
End Sub

Private Sub cmdSave_Click()
If Not myreplace Then Exit Sub
    
Inform " „ «·Õ›Ÿ"

If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub cmdStore_Click()
StoreLookup Me, oSearchStore, "CODE IN(SELECT STORE FROM FILE6_20H WHERE INV_TYPE >= 2)", , IIf(cmdStore.Tag = "", "", "ﬂ· «·„Œ«“‰")
End Sub

Private Sub cmdStoreLookup_Click()
'StoreLookup Me, oSearchStoreLook, "CODE IN(SELECT STORE FROM FILE6_20H WHERE INV_TYPE = 2) AND ISSTOP = 0"
StoreLookup Me, oSearchStoreLook, "ISSTOP = 0"
End Sub

Private Sub cmdTaxInv_Click()
Me.MousePointer = vbHourglass
Dim bDone As Boolean
bDone = TransToTaxInv
Me.MousePointer = vbNormal
If bDone Then
    MsgBox " „  ÕÊÌ· «·›« Ê—… »‰Ã«Õ"
End If
End Sub

Private Sub CmdUndo_Click()
myUndo
End Sub

Private Sub DAYALL1_Click()
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then
        myDefine
    End If
End If
End Sub

Private Sub Form_Activate()
If Not validOpen Then
    Unload Me
    Exit Sub
End If
If Not bAct Then
    If xDoc_No.Tag = LoadMode Then
        On Error Resume Next
        grid1.SetFocus
        CellPos 13, grid1.Rows - 2, grid1.Cols - 1
    ElseIf xStore.Enabled Then
       xStore.SetFocus
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
On Error GoTo myerror
OpenCon con
nRound = 0

HandleInit
Set xMan.RowSource = cmd("SELECT * FROM FILE6_25  WHERE isstop = 0  ORDER BY FILE6_25.DESCA ", con).Execute
xMan.ListField = "Desca"
xMan.BoundColumn = "Code"

Set data1.RecordSet = cmd("SELECT * FROM FILE0_40 WHERE ISSTOP = 0", con).Execute
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"
xStore.BoundText = cBranchStore

Set data2.RecordSet = mycmd(YearsString(2023), con)
Set xYear.RowSource = data2
xYear.ListField = "CODE"
xYear.BoundColumn = "CODE"


Set DATA3.RecordSet = mycmd(MonthString(12), con)
Set xMonth.RowSource = DATA3
xMonth.ListField = "CODE"
xMonth.BoundColumn = "CODE"

Set xLogo.RowSource = myRs("SELECT * FROM LOGO_CODES", con)
xLogo.ListField = "Desca"
xLogo.BoundColumn = "Code"

Set rsInv_Type = createRs(cmd("SELECT * FROM INV_TYPE_CODES", con).Execute)

Set grid1.DataSource = DATA11
Set grid2.DataSource = data12
If Not openCardTable Then myDefine
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Form_Resize()
Dim nHeight As Long
nHeight = Me.Height - (grid1.Top + 2650)
If nHeight > 3000 Then
    grid1.Height = nHeight
    grid2.Height = nHeight
    grdError.Height = nHeight
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

CardTable.Close
Set CardTable = Nothing

closeCon con

Set sales_abdfrm = Nothing
Err.Clear
End Sub

Public Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
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
        grid1.TextMatrix(Row, 15) = fnBalance(grid1.TextMatrix(Row, 1), con, xStore.BoundText) & ""
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
'Err.Clear
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
Private Sub grid1_LostFocus()
'ImgX1.Images.Clear
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
        If optType(2).Value And Val(.EditText) < 0 Then
            MsgBox "«·ﬁÌ„… «·”«·Ì… €Ì— „”„ÊÕ »Â« ·«–Ê‰ «·„»Ì⁄« "
            Cancel = True
        ElseIf optType(3).Value And Val(.EditText) >= 0 Then
            MsgBox "«·ﬁÌ„… «·„ÊÃ»… €Ì— „”„ÊÕ »Â« ·«–Ê‰ «·„— Ã⁄« "
            Cancel = True
        ElseIf Val(xinv_type.Caption) = 3 And Val(.EditText) < 0 Then
            MsgBox "«·ﬁÌ„… «·”«·»… €Ì— „”„ÊÕ »Â« ··›Ê« Ì— «·÷—Ì»Ì…"
            Cancel = True
        ElseIf Val(xinv_type.Caption) = 4 And Val(.EditText) >= 0 Then
            MsgBox "«·ﬁÌ„…«·„ÊÃ»… €Ì— „”„ÊÕ »Â« ··„— Ã⁄«  «·÷—Ì»Ì…"
            Cancel = True
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
If grid1.Rows = 2 Then Exit Sub
doprint
End Sub
Private Sub optclosed_Click(index As Integer)
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then
        myDefine
    End If
End If
End Sub

Private Sub lblAlert_Click()

End Sub

Private Sub optPrice_Click(index As Integer)
If bIg Then Exit Sub
If MsgBox(" ÕÊÌ· «·Ì " & IIf(index = 0, "”⁄— Ã„·…", "”⁄— „” Â·ﬂ"), vbOKCancel + vbDefaultButton2) <> vbOK Then
    bIg = True
    optPrice(index).Value = False
    optPrice(IIf(index = 1, 0, 1)) = True
    optPrice(IIf(index = 1, 0, 1)).SetFocus
    bIg = False
Else
    Me.MousePointer = vbHourglass
    On Error GoTo myerror
    con.Execute "UPDATE FILE6_20  SET " & _
              " FILE6_20.PRICE_C = FILE1_10." & IIf(optPrice(0).Value, "COST2", "PRICE") & "," & _
              " FILE6_20.PRICE = FILE1_10." & IIf(optPrice(0).Value, "COST2", "PRICE") & _
              " FROM FILE6_20 " & _
              " INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM" & _
              " WHERE FILE6_20.DOC_NO = " & MyParn(xDoc_No.text)
    Me.MousePointer = vbNormal
    Inform " „  ⁄œÌ· «·«”⁄«— »‰Õ«Õ"
    myLoadGrd
    CellPos 13, grid1.Rows - 2, grid1.Cols - 1
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
Me.MousePointer = vbNormal
End Sub

Private Sub optType_Click(index As Integer)
If bIg Then Exit Sub
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then
        myDefine
    End If
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

Private Sub xDiscount_Change()
If bIg Then Exit Sub
CalcTotals
End Sub

Private Sub xDiscount_LostFocus()
myLostFocus xDiscount
'CalcTotals
End Sub
Private Function myValid(Optional bIgMsg As Boolean = False, Optional bIgMan As Boolean = False) As Boolean
'If xDoc_No.Text = "" Then
'    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
'    Exit Function
'End If

If Trim(xCode.text) = "" Then
    MsgBox "·« ÌÊÃœ ﬂÊœ ⁄„Ì· «Ê ﬂÊœ „Ê—œ"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If Not xStore.MatchedWithList Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
    Exit Function
End If

If xCodeDesca.Caption = "" Then
    If Not bIgMsg Then MsgBox "·« ÌÊÃœ ﬂÊœ ⁄„Ì·"
    Exit Function
End If

If Val(xinv_type.Caption) < 2 Then
    If Not bIgMsg Then MsgBox "·« ÌÊÃœ ‰Ê⁄ ›« Ê—…"
    Exit Function
End If
If Val(xinv_type.Caption) <> 2 Then
    If (Not ValidQuant) Then
        MsgBox "ﬂ„Ì… „ÊÃ»… Êﬂ„Ì… ”«·»… ›Ï ‰›” «·›« Ê—…"
        Exit Function
    End If
End If
myValid = True
End Function
Private Sub myload()
XBRANCH.Caption = CardTable!branch & ""
xDoc_No.text = CardTable!doc_no
xDate_due.text = myFormat_p(CardTable!DATE_DUE)
xDoc_no2.Caption = CardTable!Doc_no2
xdoc_no_tax.Caption = Mid(CardTable!doc_no_Tax & "", 7, 6)
xdoc_no_tax.Tag = CardTable!doc_no_Tax & ""
xdoc_no_flag.Caption = CardTable!DOC_NO_FLAG & ""
XISINVOICE.Value = IIf(CardTable!ISINVOICE, 1, 0)
xIs_collect.Value = IIf(CardTable!IS_COLLECT, 1, 0)
xLogo.BoundText = CardTable!logo & ""
xMan.BoundText = CardTable!MAN & ""
xPrice_list.Tag = CardTable!PRICE_LIST & ""
xPrice_list.Caption = CardTable!LIST_DESCA & ""

XDATESALES1.text = myFormat_p(CardTable!DATESALES1)
XDATESALES2.text = myFormat_p(CardTable!DATESALES2)
xpo_no.text = CardTable!po_no & ""
xinv_no.text = CardTable!inv_no & ""
xinv_type.Caption = CardTable!INV_TYPE & ""
xDate.text = myFormat_p(CardTable!Date)
xStore.BoundText = CardTable!STORE & ""
xNotes.text = CardTable!notes & ""
xCode.text = CardTable!code & ""
xCodeDesca.Caption = CardTable!ClientDesca & ""
panel1(2).Caption = CardTable!UserName & ""
XINV_TAX.Caption = CardTable!inv_tax & ""

bIg = True
xDiscount.text = Myvalue(CardTable!discount)
xTax.text = Myvalue(CardTable!tax)

If Val(xinv_type.Caption) = 2 And chkTotalInv.Value = 1 Then
    chkTotalInv.Value = 0
    grid1.Visible = True
    grid2.Visible = False
End If

bIg = False
panel1(1).Caption = CardTable!USER_IP & ""
Handlecontrols LoadMode

If chkTotalInv.Value = 0 Then
    myLoadGrd
    CellPos 13, grid1.Rows - 2, grid1.Cols - 1
ElseIf chkTotalInv.Value = 1 Then
    myloadgrd2
End If
Exit Sub
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub myDefine()
xDoc_No.text = addNewDoc
xDoc_no2.Caption = Mid(xDoc_No.text, 7, 6)
xDate_due.text = ""
XINV_TAX.Caption = ""
xMan.BoundText = ""
xLogo.BoundText = "1"
xdoc_no_tax.Caption = ""
xdoc_no_flag.Caption = ""
XDATESALES1.text = ""
XDATESALES2.text = ""
xIs_collect.Value = 0

xPrice_list.Tag = ""
xPrice_list.Caption = ""

xpo_no.text = ""

If optType(4).Value Or optType(5).Value Or optType(6).Value Then
    XISINVOICE.Value = 1
Else
    XISINVOICE.Value = 0
End If

If optType(1).Value Or optType(2).Value Or optType(3).Value Then
    xinv_type.Caption = "2"
ElseIf optType(5).Value Then
    xinv_type.Caption = "3"
ElseIf optType(6).Value Then
    xinv_type.Caption = "4"
Else
    xinv_type.Caption = ""
End If


'xDoc_No.Tag = DefineMode
'xDoc_no2.text = ""

'xinv_no.text = ""
'xMan.BoundText = "0001"
'xusername.Caption = cusername
xStore.BoundText = ""
xCode.text = ""
xCodeDesca.Caption = ""

xDate.text = myFormat_p(Date)


xTotalQuant.Caption = ""
xTotalItem.Caption = ""

bIg = True
xDiscount.text = ""
xRate.text = ""
bIg = False
xTotal.Caption = ""
xNotes.text = ""

panel1(1).Caption = GetComputerName
'xusername.Caption = cusername

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
If optType(0).Value Or optType(4).Value Then
    cmdNewInv.Enabled = False
Else
    cmdNewInv.Enabled = bEdit And nMode = LoadMode
End If

xDoc_No.Tag = nMode


bEditRecord = bEdit And XINV_TAX.Caption = "" And xdoc_no_tax.Caption = ""
bEditRecord = bEditRecord And (xIs_collect.Value = 0 Or nMode = LoadMode)
bEditRecord = xinv_type.Caption <> ""

CMD_ADD_RECEVD.Enabled = nMode = DefineMode And bEditRecord And (optType(2).Value Or optType(5).Value)


fmCollectTrans.Visible = xIs_collect.Value = 1

cmdSign.Enabled = XINV_TAX.Caption = ""
cmdSignGroup.Enabled = XINV_TAX.Caption = ""
cmdTaxInv.Enabled = xdoc_no_tax.Caption = ""

xStore.Enabled = xIs_collect.Value = 0
cmdStoreLookup.Enabled = xIs_collect.Value = 0

cmdSave.Enabled = bEditRecord
cmddel.Enabled = (bEditRecord Or bDelRecord) And nMode = LoadMode
fmPrice.Enabled = bEditRecord

xDiscount.Enabled = bEditRecord And Val(xinv_type.Caption) = 2 And xIs_collect.Value = 0
xRate.Enabled = bEditRecord And Val(xinv_type.Caption) = 2 And xIs_collect.Value = 0
xRate_Tax.Enabled = bEditRecord And Val(xinv_type.Caption) > 2 And xIs_collect.Value = 0
cmdPurchase.Enabled = Val(xinv_type.Caption) = 2 And xIs_collect.Value = 0

cmdCSV.Enabled = bEditRecord And nMode = LoadMode

fmTransTax.Visible = Val(xinv_type.Caption) = 2 And nMode = LoadMode
fmShowTotal.Visible = Val(xinv_type.Caption) > 2 And nMode = LoadMode
fmE_Invoice.Visible = Val(xinv_type.Caption) > 2 And nMode = LoadMode

If grdError.Visible Then grdError.Visible = False

Dim nRecord As Long, nRecords As Long
'retRecords xdoc_no_flag.Caption, nRecords, nRecord
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
If optType(6).Value Or optType(3).Value Or Val(xinv_type.Caption) = 4 Then
    grid1.TextMatrix(Row, 10) = "-1"
Else
    grid1.TextMatrix(Row, 10) = "1"
End If
If xPrice_list.Tag <> "" Then
    Dim loctable As New ADODB.RecordSet
    Set loctable = mycmd("Select price " & _
                    " from file6_53  " & _
                     " where doc_no = " & xPrice_list.Tag & _
                     " and item = " & MyParn(cm.Parameters("@ITEM").Value), con)
    If loctable.EOF Then
        grid1.TextMatrix(Row, 11) = IIf(optPrice(0).Value, cm.Parameters("@COST2").Value, cm.Parameters("@PRICE").Value)
    Else
        grid1.TextMatrix(Row, 11) = loctable!price
    End If
    Set loctable = Nothing
Else
    grid1.TextMatrix(Row, 11) = IIf(optPrice(0).Value, cm.Parameters("@COST2").Value, cm.Parameters("@PRICE").Value)
End If
grid1.TextMatrix(Row, 14) = cm.Parameters("@COST").Value
grid1.TextMatrix(Row, 15) = fnBalance(cm.Parameters("@ITEM").Value, con, xStore.BoundText) & ""
GrdDesc = True
End With
End Function
Private Function CalcTotals(Optional nMode As Integer = 0)
Dim nTotalQuant As Double, nTotalQuant2 As Double, nTotalItem As Double, nTotalItemNoDiscount As Double, nTotalDiscount As Double, nTotalMin As Double
Dim nDisountRow As Double
With grid1
Dim i As Long
Dim nPrice_C As Long
For i = 1 To grid1.Rows - 2
    .TextMatrix(i, 13) = .ValueMatrix(i, 10) * .ValueMatrix(i, 11)
    nTotalQuant = nTotalQuant + grid1.ValueMatrix(i, 10)
    nTotalItem = mRound(nTotalItem + (.ValueMatrix(i, 10) * .ValueMatrix(i, 11)), 6)
Next
xTotalQuant.Caption = Myvalue(nTotalQuant)
xTotalItem.Caption = Myvalue(mRound(nTotalItem))


bIg = True
If Val(xTotalItem.Caption) <> 0 Then
    If Round(Val(xRate.text), nRound) <> Round(Val(xDiscount.text) / Val(xTotalItem.Caption) * 100, nRound) Then
        xRate.text = Myvalue(Round((Val(xDiscount.text) / Val(xTotalItem.Caption)) * 100, nRound))
    End If
Else
    xRate.text = ""
End If


If IsDate(XDATESALES1.text) Or IsDate(XDATESALES2.text) Then
    xRate_Tax.text = "14"
    xTax.text = mRound(Val(xTotalItem.Caption) * 0.14)
Else
    If (Val(xTotalItem.Caption) - Val(xDiscount.text)) <> 0 Then
        If Round(Val(xRate_Tax.text), nRound) <> Round(Val(xTax.text) / (Val(xTotalItem.Caption) - Val(xDiscount.text)) * 100, nRound) Then
            xRate_Tax.text = Myvalue(Round((Val(xTax.text) / (Val(xTotalItem.Caption) - Val(xDiscount.text))) * 100, nRound))
        End If
    Else
        xRate_Tax.text = ""
    End If
End If
For i = 1 To .Rows - 2
    If .ValueMatrix(i, 9) > 0 Then
         nPrice = .ValueMatrix(i, 9) / ((1 + (Val(xRate_Tax.text) / 100)))
        .TextMatrix(i, 12) = Myvalue(mRound((nPrice - .ValueMatrix(i, 11)) / nPrice, 4) * 100)
    Else
        .TextMatrix(i, 12) = ""
    End If
Next

xTotal.Caption = mRound((nTotalItem - Val(xDiscount.text)) + Val(xTax.text), 2)
bIg = False
End With
End Function
Private Sub CardLookup(Optional pFilter As String = "")
Dim Generalarray(5)
Dim listarray(3, 5)
Dim GrdArray(7, 1)
Set Generalarray(0) = Me

Generalarray(1) = "SELECT TOP 1000 " & _
                  " FILE6_20H.DOC_NO ," & _
                  " FILE6_20H.DOC_NO2," & _
                  " FORMAT(FILE6_20H.date,'yyyy/M/d')," & _
                  " FILE0_50.DESCA," & _
                  " FILE3_10.DESCA," & _
                  " FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT," & _
                  " INV_TYPE_CODES.DESCA," & _
                  " FILE6_20H.NOTES" & _
                  " FROM FILE6_20H" & _
                  " INNER JOIN FILE3_10 ON FILE6_20H.CODE = FILE3_10.CODE" & _
                  " INNER JOIN INV_TYPE_CODES ON FILE6_20H.INV_TYPE = INV_TYPE_CODES.CODE" & _
                  " INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE" & _
                  " LEFT JOIN FILE0_50 ON FILE6_20H.BOX = FILE0_50.CODE"

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & " WHERE " & cFilter
End If

Generalarray(2) = "Order by DATE DESC,doc_no2 DESC "
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ «·„” ‰œ"
listarray(0, 1) = "(FILE6_20H.DOC_NO = 'cFilter' OR FILE6_20H.DOC_NO2 = 'cFilter')"

listarray(1, 0) = "«· «—ÌŒ"
listarray(1, 1) = "(##FILE6_20H.DATE##)"

listarray(2, 0) = "«·⁄„Ì·"
listarray(2, 1) = "(%%FILE3_10.DESCA%%)"

listarray(3, 0) = "«·»Ì«‰"
listarray(3, 1) = "(%%FILE6_20H.NOTES%%)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(1, 1) = 1200

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1300

GrdArray(3, 0) = "«·Õ“‰…"
GrdArray(3, 1) = 2000

GrdArray(4, 0) = "«·⁄„Ì·"
GrdArray(4, 1) = 3500

GrdArray(5, 0) = "«·ﬁÌ„…"
GrdArray(5, 1) = 1300

GrdArray(6, 0) = "‰Ê⁄ «·›« Ê—…"
GrdArray(6, 1) = 1500

GrdArray(7, 0) = "«·»Ì«‰"
GrdArray(7, 1) = 3500

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
    con.Execute " UPDATE FILE6_20H SET ISRET = " & xIsRet.Value & " , ISNEW = 1 WHERE DOC_NO = " & MyParn(xDoc_No.text), nRec
    If nRec = 1 Then MsgBox "  „  ⁄œÌ· «·„” ‰œ "
End If
End Sub
Private Sub xMAN_GotFocus()
myGotFocus xMan
End Sub
Private Sub xMan_LostFocus()
myLostFocus xMan
If xMan.MatchedWithList Then Exit Sub
If Not IsDgt(Trim(xMan.text)) Then xMan.BoundText = ""
xMan.BoundText = RetZero(Trim(xMan.text), 4)
If Not xMan.MatchedWithList Then xMan.BoundText = ""
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
Set Phonesfrm.myForm = Me
Set Phonesfrm.con = con
Phonesfrm.Show 1
End Sub

Private Sub xPrinted_Click()
If bIg Then Exit Sub
'If xPrinted.Value = 1 Then Exit Sub
con.Execute "update file6_20h set printed = " & xPrinted.Value & ",closed = " & xPrinted.Value & " from file6_20h where doc_no = " & MyParn(xDoc_No.text)
MsgBox " „ › Õ «·›« Ê—… »‰Ã«Õ"
If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
    If Not openCardTable Then
        myDefine
    End If
End If
End Sub

Private Sub xinv_type_Change()
If xinv_type.Caption <> "" Then
'    rsInv_Type.Find "CODE = " & xinv_type.Caption, , adSearchBackward, adBookmarkFirst
'    If Not rsInv_Type.EOF Then
'        lbl_inv_type.Caption = rsInv_Type!desca
'    Else
'        lbl_inv_type.Caption = ""
'    End If
    lblInv_type.Caption = rsField(rsInv_Type, xinv_type.Caption, "desca") & ""
Else
    lblInv_type.Caption = ""
End If
End Sub

Private Sub xPO_NO_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then OrderNoLook
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

Private Sub xPrice_list_Click()
PriceListLook Me, oSearchList, , , IIf(xPrice_list.Tag <> "", "»œÊ‰ ﬁ«∆„… ”⁄—", "")
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

Private Sub xRate_LostFocus()
myLostFocus xRate
'If Val(xTotalItem.Caption) <> 0 Then
'    If Round(Val(xRate.text), nRound) <> Round(Val(xDiscount.text) / Val(xTotalItem.Caption) * 100, nRound) Or xDiscount.Locked Then
'        xDiscount.text = Round((Val(xRate.text) * Val(xTotalItem.Caption)) / 100, 0)
'    End If
'Else
'    xDiscount.text = ""
'End If
'CalcTotals
End Sub
Private Function RetItemBalance(cItem, cStore, dDate) As Double
If cItem = "" Then Exit Function
movetable.Seek Array(cItem, cStore), adSeekFirstEQ
Do Until movetable.EOF
    If IsNull(movetable!Date) Then Exit Do
    If Trim(movetable!Item) <> cItem Or cStore <> movetable!STORE Or DateValue(movetable!Date) > DateValue(Format(dDate, "dd-mm-yyyy")) Then Exit Do
    RetItemBalance = RetItemBalance + TurnValue(movetable!In, Null, 0) - TurnValue(movetable!out, Null, 0)
    movetable.MoveNext
Loop
End Function
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Sub fixGrd()
With grid1
'                 0         1          2         3         4          5             6          7          8         9          10        11           12           13          14            15
.FormatString = "„.|" & "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "”⁄—|" & "⁄œœ|" & "”⁄— «·»Ì⁄|" & "‰”»… Œ’„|" & "«·≈Ã„«·Ï|" & "«· ﬂ·›…|" & "«·—’Ìœ|"
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

If IsDate(XDATESALES1.text) Or IsDate(XDATESALES2.text) Then
    grid1.BackColorFixed = &HC0E0FF
Else
    grid1.BackColorFixed = &HE0E0E0
End If
End With
End Sub
Private Sub fixGrdError()
With grdError
.FormatString = "—ﬁ„ «·”ÿ—|" & "»«—ﬂÊœ|" & "«·Œÿ√"
.ColAlignment(0) = flexAlignRightCenter
.ColAlignment(1) = flexAlignRightCenter
.ColAlignment(2) = flexAlignRightCenter
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter

.ColWidth(0) = 1300
.ColWidth(1) = 2000
.ColWidth(2) = 9000
End With
End Sub

Private Sub xRateDis_Lostfocus()
xDiscount.text = Fix((Val(xTotalItem.Caption) * Val(xRateDis.text) / 100))
End Sub

Private Sub XSALES_RET_Change()
XSALES_RET.Visible = XSALES_RET.Caption <> ""
End Sub

Private Sub xSendRc_Click()
If bIgClick Then Exit Sub
If xDoc_No.text <> "" Then
    On Error GoTo myerror
    If xUUID_RC.Caption <> "" Then
        If Not IsEmpty(myField("select doc_no from file6_20h where PREVIOUS_UUID = " & MyParn(xUUID_RC.Caption), con)) Then
            MsgBox "—ﬁ„ „”·”· „—Ã⁄Ì ”«»ﬁ ··›« Ê—… —ﬁ„ " & MyParn(xDoc_No.text)
            'xSendRc.Value = IIf(xSendRc.Value = 1, 0, 1)
        End If
    End If
    con.Execute "UPDATE FILE6_20H SET FILE6_20H.sendRc = " & xSendRc.Value & " FROM FILE6_20H WHERE FILE6_20H.DOC_NO = " & MyParn(xDoc_No.text)
    Inform " „ «· ⁄œÌ· »‰Ã«Õ"
End If
Finaly:
Handlecontrols xDoc_No.Tag
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
bIgClick = True
xSendRc.Value = IIf(xSendRc.Value = 1, 0, 1)
bIgClick = False
GoTo Finaly
End Sub

Private Sub xRate_Tax_Change()
If bIg Then Exit Sub
bIg = True
If (Val(xTotalItem.Caption) - Val(xDiscount.text)) <> 0 Then
    If Round(Val(xRate_Tax.text), 0) <> Round(Val(xTax.text) / (Val(xTotalItem.Caption) - Val(xDiscount.text)) * 100, 0) Then
        xTax.text = Round((Val(xRate_Tax.text) * (Val(xTotalItem.Caption) - Val(xDiscount.text))) / 100, 2)
    End If
Else
    xTax.text = ""
End If
CalcTotals
bIg = False
End Sub

Private Sub xStore_Validate(Cancel As Boolean)
    If Trim(xStore.BoundText) = "" Then Cancel = True
End Sub
Private Sub xTax_LostFocus()
If bIg Then Exit Sub
CalcTotals
End Sub
Private Function RemoveItem(Row As Long) As Boolean
If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel) <> vbOK Then
    Exit Function
End If

If grid1.TextMatrix(Row, grid1.Cols - 1) <> "" Then
    On Error GoTo myerror
    con.Execute "Delete  From FILE6_20 where id = " & grid1.TextMatrix(Row, grid1.Cols - 1)
End If
grid1.RemoveItem Row
RemoveItem = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Sub myproc2(nDoc_no)
xDoc_No = nDoc_no
myUndo
'CardTable.Find "Doc_no = " & MyParn(nDoc_no), , adSearchForward, adBookmarkFirst
'If Not CardTable.EOF Then
'    myload
'Else
'    MsgBox "—ﬁ„ «·›« Ê—… €Ì— ’ÕÌÕ"
'    Unload Me
'End If
End Sub
Function itemPrice(cItem) As Single
    itemPrice = GetDesca("select PRICE from file1_10 where item = " & MyParn(cItem), con)
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
        aInsert = AddFlag(aInsert, "MAN", addstring("0001"))
        aInsert = AddFlag(aInsert, "USER_IP", addstring(cIpName))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.text))
            con.Execute addInsert(aInsert, "FILE6_20")
        Else
            con.Execute addUpdate(aInsert, "FILE6_20", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Function myreplaceGrdReplace() As Boolean
Dim aInsert As Variant
Dim cString As String
cString = "select Max(Offer_no) " & _
          " from file6_20 " & _
          " where doc_no = " & MyParn(xDoc_No.text)
Dim nOffer As Long
If Val(frmOfferReplace.XQTY2.Caption) > 0 Then
    nOffer = Val(myField(cString, con) & "") + 1
End If
                                    
With frmOfferReplace.grid1
    For i = 1 To .Rows - 2
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.text))
        aInsert = AddFlag(aInsert, "ITEM", addstring(.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "QUANT", .ValueMatrix(i, 10))
        aInsert = AddFlag(aInsert, "PRICE", .ValueMatrix(i, 9))
        aInsert = AddFlag(aInsert, "PRICE_c", .ValueMatrix(i, 14))
        aInsert = AddFlag(aInsert, "COST", .ValueMatrix(i, 13))
        If frmOfferReplace.xMan.Caption <> "" Then
            aInsert = AddFlag(aInsert, "MAN", addstring(frmOfferReplace.xMan.Caption))
        ElseIf xMan.MatchedWithList Then
            aInsert = AddFlag(aInsert, "MAN", addstring(xMan.BoundText))
        End If
        aInsert = AddFlag(aInsert, "OFFER_NO", nOffer)
        aInsert = AddFlag(aInsert, "DOC_OFFER", addstring(frmOfferReplace.xDoc_No.Caption))
        aInsert = AddFlag(aInsert, "DISCOUNT_OFFER", Val(frmOfferReplace.xDiscount.Caption))
        aInsert = AddFlag(aInsert, "DISCOUNT_OFFER_RATE", Val(frmOfferReplace.xRate.Caption))
        aInsert = AddFlag(aInsert, "USER_IP", addstring(cIpName))
        con.Execute addInsert(aInsert, "FILE6_20")
    Next
End With
myreplaceGrdReplace = True
End Function
Private Function myreplaceGrdOffer() As Boolean
Dim aInsert As Variant
Dim cString As String
cString = "select Max(Offer_no) " & _
          " from file6_20 " & _
          " where doc_no = " & MyParn(xDoc_No.text)
Dim nOffer As Long
nOffer = Val(myField(cString, con) & "") + 1
                                    
With frmOffer.grid1
    For i = 1 To frmOffer.grid1.Rows - 2
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.text))
        aInsert = AddFlag(aInsert, "ITEM", addstring(.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "QUANT", .ValueMatrix(i, 10))
        aInsert = AddFlag(aInsert, "PRICE", .ValueMatrix(i, 9))
        aInsert = AddFlag(aInsert, "PRICE_c", .ValueMatrix(i, 9))
        aInsert = AddFlag(aInsert, "COST", .ValueMatrix(i, 13))
        aInsert = AddFlag(aInsert, "MAN", addstring(xMan.BoundText))
        aInsert = AddFlag(aInsert, "OFFER_NO", nOffer)
        aInsert = AddFlag(aInsert, "DOC_OFFER", addstring(frmOffer.xDoc_No.Caption))
        aInsert = AddFlag(aInsert, "DISCOUNT_OFFER", Val(frmOffer.xDiscount.Caption))
        aInsert = AddFlag(aInsert, "DISCOUNT_OFFER_RATE", Val(frmOffer.xRate.Caption))
        aInsert = AddFlag(aInsert, "USER_IP", addstring(cIpName))
        con.Execute addInsert(aInsert, "FILE6_20")
    Next
End With
myreplaceGrdOffer = True
End Function
Private Sub myLoadGrdOld()
Dim cString As String
cString = "SELECT FILE6_20.ITEM," & _
          "FILE1_10.MOSM," & _
          "FACT.DESCA," & _
          "FILE1_10.SUPP," & _
          "FILE1_10.MODELFACT0," & _
          "FILE1_10.DESCA," & _
          "FILE1_10.SCAL," & _
          "FILE1_10.COLOR," & _
          "FILE6_20.QUANT," & _
          "FILE6_20.PRICE," & _
          "FILE6_20.TOTAL," & _
          "FILE6_20.COST," & _
          "BALANCE.BALANCE," & _
          "FILE6_20.ID " & _
          " FROM FILE6_20 " & _
          " INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM" & _
          " LEFT JOIN (" & _
          "SELECT ITEM,SUM([IN] - OUT) AS BALANCE FROM FILE1_11_ALL GROUP BY ITEM)" & _
          " AS BALANCE ON FILE6_20.ITEM = BALANCE.ITEM" & _
          " LEFT JOIN FACT ON FACT.CODE = FILE1_10.FACT"
cString = cString & " WHERE FILE6_20.DOC_NO = " & MyParn(xDoc_No.text)
cString = cString & " ORDER BY FILE6_20.ID"
Set DATA11.RecordSet = cmd(cString, con).Execute
myAddItem
CalcTotals
fixGrd
End Sub
Private Sub myLoadGrd()
Dim aPrm As Variant
If xIs_collect.Value = 1 Then
    aPrm = AddFlag(aPrm, "INV_NO", xDoc_No.text)
    Set DATA11.RecordSet = cmd("[dbo].[sp_invoice_whole_collect]", con, adStoredProc, aPrm).Execute
Else
    aPrm = AddFlag(aPrm, "DOC_NO", xDoc_No.text)
    aPrm = AddFlag(aPrm, "STORE", xStore.BoundText)
    If chkNoBalance.Value = 0 Then
        Set DATA11.RecordSet = cmd("[dbo].[sp_invoice_whole]", con, adStoredProc, aPrm).Execute
    Else
        Set DATA11.RecordSet = cmd("[dbo].[sp_invoice_whole_no_balance]", con, adStoredProc, aPrm).Execute
    End If
End If
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
If (optType(2).Value Or Val(xinv_type.Caption) = 3) And .ValueMatrix(Row, 10) < 0 Then
    Exit Function
ElseIf (optType(3).Value Or Val(xinv_type.Caption) = 4) And .ValueMatrix(Row, 10) >= 0 Then
    Exit Function
End If
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
Private Function retRecords(pDoc_no, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New ADODB.RecordSet
If pDoc_no <> "" Then
    'cString = "SELECT Count(FILE6_20H.DOC_NO) AS records,COUNT(CASE WHEN FILE6_20H.DOC_NO_FLAG <= " & MyParn(pDoc_no) & " THEN 1 END) AS record"
    cString = "SELECT Count(FILE6_20H.DOC_NO) AS records,COUNT(CASE WHEN FILE6_20H.DOC_NO <= " & MyParn(pDoc_no) & " THEN 1 END) AS record"
Else
    cString = "SELECT Count(FILE6_20H.DOC_NO) AS records,0 as record"
End If

cString = cString & " FROM FILE6_20H " & _
                    " INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE"

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
retFilter = "FILE6_20H.INV_TYPE > 1"
If optType(1).Value Then
    retFilter = retFilter & " AND FILE6_20H.INV_TYPE = 2"
ElseIf optType(2).Value Then
    retFilter = retFilter & " AND FILE6_20H.INV_TYPE = 2 AND FILE6_20H.IS_RETURN = 0"
ElseIf optType(3).Value Then
    retFilter = retFilter & " AND FILE6_20H.INV_TYPE = 2 AND FILE6_20H.IS_RETURN = 1"
ElseIf optType(4).Value Then
    retFilter = retFilter & " AND FILE6_20H.INV_TYPE > 2"
ElseIf optType(5).Value Then
    retFilter = retFilter & " AND FILE6_20H.INV_TYPE = 3"
ElseIf optType(6).Value Then
    retFilter = retFilter & " AND FILE6_20H.INV_TYPE = 4"
End If

If cmdStore.Tag <> "" Then
    retFilter = retFilter & " AND FILE6_20H.STORE = " & MyParn(cmdStore.Tag)
End If

If cmdClient.Tag <> "" Then
    retFilter = retFilter & " AND FILE6_20H.CODE = " & MyParn(cmdClient.Tag)
End If

If chkTotal.Value = 1 Then
    retFilter = retFilter & " AND FILE6_20H.IS_COLLECT = 1"
End If

If chkNoTax.Value = 1 Then
    retFilter = retFilter & " AND FILE6_20H.ISINVOICE = 1 AND FILE6_20H.INV_TAX IS NULL"
End If

If IsDate(xdate_filter.text) Then
    retFilter = retFilter & Tr(retFilter) & "FILE6_20H.DATE = " & DateSq(xdate_filter.text)
ElseIf chkDay.Value = 1 Then
    retFilter = retFilter & Tr(retFilter) & "FILE6_20H.DATE = " & DateSq(Date)
Else
    If xMonth.BoundText <> "" And xYear.BoundText <> "" Then
        retFilter = retFilter & Tr(retFilter) & "FILE6_20H.YEAR_MONTH = " & MyParn(xYear.BoundText & "-" & RetZero(xMonth.BoundText, 2))
    ElseIf xYear.BoundText <> "" Then
        retFilter = retFilter & Tr(retFilter) & "FILE6_20H.YEAR_MONTH LIKE " & MyParn(xYear.BoundText & "%")
    End If
End If

If sDoc_no <> "" Then
    retFilter = "DOC_NO = " & MyParn(sDoc_no)
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

If xPrice_list.Tag = "" Then
    Set loctable = cmd("select doc_no,notes from file6_53h where code = " & MyParn(xCode.text), con).Execute
    If Not loctable.EOF Then
        xPrice_list.Caption = loctable!notes & ""
        xPrice_list.Tag = loctable!doc_no & ""
    End If
End If
Set loctable = Nothing
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
Private Sub HandleInit()
bEdit = True
fixGrdError
End Sub
Private Sub xTotalItemDiscount_Change()
xTotalItem.Visible = Val(xTotalItemDiscount) <> 0
lblTotalitem.Visible = Val(xTotalItemDiscount) <> 0
End Sub
Private Sub xTotalQuant_Change()
xTotalQuant.ForeColor = IIf(Val(xTotalQuant.Caption) < 0, vbRed, vbBlack)
lblQuant.ForeColor = IIf(Val(xTotalQuant.Caption) < 0, vbRed, vbBlack)
End Sub
Private Sub doprint()
Dim temptable As New ADODB.RecordSet
Dim sourcetable As New ADODB.RecordSet
Dim lPrice As Boolean
Dim nFBal As Double
Dim nPay As Double
Dim nBal As Double
' ???? ????? ??? ?????? ??????


Dim con_image As New ADODB.Connection
openConDoc con_image
 
'picPrint.Picture = LoadPicture("")
imgxPrint.Images.Clear
 
 
'Picture3.Picture = LoadPicture(RetPhoto(xCode.text))
'aRet = retDimImg(Picture3.Width, Picture3.Height, 2500 / 15, 2500 / 15)

'Set imgx2.Picture = LoadPicture(RetPhoto(xCode.text))
'imgx2.Effects.Resize Val(retFlag(aRet, "width")), Val(retFlag(aRet, "height"))
 
 
Dim nWidth As Long, nHeight As Long
Dim imageTable As New ADODB.RecordSet
If xLogo.MatchedWithList Then
    Set imageTable = cmd("select * from logo_codes where code = " & IIf(xLogo.MatchedWithList, xLogo.BoundText, 1), con).Execute
    
    If Not imageTable.EOF Then
        If Not IsNull(imageTable!Image) Then

            Set picPrint.DataSource = imageTable
            picPrint.DataField = "image"

            GetNewSize picPrint.Width, picPrint.Height, 500, nHeight
            aRet = retDimImg(picPrint.Width, picPrint.Height, 600, 200)

             imgxPrint.Import.FromMemoryFile (imageTable!Image)
            imgxPrint.Effects.Resize 500, nHeight
        End If
    End If
End If

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable


Dim loctable As New ADODB.RecordSet, cString As String
cString = "SELECT FILE6_20H.DOC_NO2," & _
          " FILE6_20H.DATE," & _
          " FILE6_20H.DISCOUNT," & _
          " FILE6_20H.DISCOUNT_RATE," & _
          " FILE6_20H.TOTAL_ITEM," & _
          " FILE3_10.DESCA AS CODE_DESCA," & _
          " FILE6_20H.NOTES," & _
          " FILE6_20.ITEM," & _
          " FILE6_20.QUANT," & _
          " FILE6_20.PRICE," & _
          " FILE1_10.DESCA AS ITEM_DESCA," & _
          " FILE6_20.TOTAL," & _
          " file1_10.modelfact," & _
          " FACT.DESCA AS FACTDESCA," & _
          " FILE1_10.COLOR," & _
          " FILE1_10.C_COLOR," & _
          " FILE1_10.SCAL," & _
          " FILE1_10.MODEL" & _
          " FROM FILE6_20 " & _
          " INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO" & _
          " INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM" & _
          " INNER JOIN FILE3_10 ON FILE6_20H.CODE = FILE3_10.CODE" & _
          " LEFT JOIN FACT ON FACT.CODE = FILE1_10.[FACT]" & _
          " WHERE  FILE6_20.DOC_NO = " & MyParn(xDoc_No.text)
Set loctable = cmd(cString, con).Execute


Do While Not loctable.EOF
    temptable.AddNew
    
    cString = " select top 1 * " & _
             " from MODEL_IMAGES where model = " & MyParn(loctable!MODEL) & _
             " AND (C_COLOR = " & loctable!c_Color & _
             " OR C_COLOR IS NULL)" & _
             " ORDER BY C_COLOR,ID"
     
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
'
'                SaveImageToRs ImgxPrint2.Picture, temptable, "picture2"
'            End If
'        End If
'    End If
    
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
    
    If imgxPrint.Images.Count > 0 Then
        SaveImageToRs imgxPrint.Picture, temptable, "picture1"
    End If
        
    temptable!STR6 = RetZero(loctable!Doc_no2, 6)
    temptable!Date1 = myFormat_p(loctable!Date)
    temptable!STR7 = loctable!CODE_desca
    temptable!str8 = loctable!notes
    
    temptable!str2 = loctable!FACTDESCA
    temptable!str1 = loctable!Item
    temptable!str13 = TurnValue(DelZero(loctable!modelfact))
    
    temptable!str3 = loctable!ITEM_DESCA
    temptable!STR4 = loctable!color
    temptable!STR5 = loctable!SCAL
    temptable!str16 = loctable!MODEL
    
    temptable!VAL3 = loctable!Quant
    temptable!VAL1 = loctable!price
    temptable!VAL4 = loctable!TOTAL
    temptable!val5 = loctable!discount
    temptable!VAL7 = loctable!TOTAL_ITEM - loctable!discount
    
    
    temptable!VAL14 = i
    temptable.Update
    loctable.MoveNext
Loop
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
REPORT1.Destination = crptToWindow
If chkImage.Value = 1 Then
    REPORT1.ReportFileName = App.Path & "\Reports\Print_SALES_images.rpt"
Else
    REPORT1.ReportFileName = App.Path & "\Reports\Print_SALES.rpt"
End If
REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub

Private Sub xYear_Change()
If bIg Then Exit Sub
If chkDay.Value = 1 Or IsDate(xdate_filter.text) Then Exit Sub
myUndo
End Sub
Private Sub xMonth_Change()
If bIg Then Exit Sub
If chkDay.Value = 1 Or IsDate(xdate_filter.text) Or Not Val(xYear.text) < 2000 Or Val(xYear.text) > 2100 Then Exit Sub
myUndo
End Sub
Private Function getCSV()
Dim cFileName As String
Dim sb As New ChilkatStringBuilder
Dim strTab As New ChilkatStringTable

Common1.InitDir = ""
Common1.FileName = ""
Common1.Filter = "CSV (*.CSV*)|*.CSV*"
Common1.ShowOpen

cFileName = Common1.FileName

If cFileName = "" Then Exit Function

Dim cSv As New ChilkatCsv
cSv.Delimiter = ","

cSv.HasColumnNames = 1
nAccess = sb.LoadFile(cFileName, "UTF-8")
If nAccess = 0 Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰  Õ„Ì· «·„·›"
    Exit Function
End If

nAccess = cSv.LoadFromString(sb.GetAsString)

If cSv.NumRows < 1 Then Exit Function

Dim Tb As New ChilkatStringBuilder
Dim cString As New ChilkatStringBuilder
Dim loctable As New ADODB.RecordSet

prog1.Visible = True

Dim sCaption As String
Dim bSkip As Boolean
Dim sitem As String
Dim nquant As Long
Dim nPrice As Double
Dim i As Long


sCaption = Me.Caption
grdError.Rows = 1
grdError.Visible = False
For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
    
    sitem = Trim(cSv.GetCell(i, 0))
    nquant = Val(cSv.GetCell(i, 1))
    nPrice = Val(cSv.GetCell(i, 2))
    
    If Trim(sitem) <> "" And Len(sitem) < 3 Or Len(sitem) > 30 Then
        grdError.AddItem ""
        grdError.TextMatrix(grdError.Rows - 1, 0) = i + 1
        grdError.TextMatrix(grdError.Rows - 1, 1) = sitem
        grdError.TextMatrix(grdError.Rows - 1, 2) = "ﬂÊœ €Ì— „ÊÃÊœ"
        If Not grdError.Visible Then grdError.Visible = True
    ElseIf Trim(sitem) = "" And nquant = 0 Then
    ElseIf (nquant = 0) Then
        '            grdError.AddItem ""
        '            grdError.TextMatrix(grdError.Rows - 1, 0) = i + 1
        '            grdError.TextMatrix(grdError.Rows - 1, 1) = sitem
        '            grdError.TextMatrix(grdError.Rows - 1, 2) = "·«  ÊÃœ ﬂ„Ì…"
    Else
        Dim cm As New ADODB.command
        Set cm = cmd("dbo.sp_item", con, adStoredProc, AddFlag(Empty, "BARCODE_FIND", sitem))
        cm.Execute
        
        If IsNull(cm.Parameters("@ITEM").Value) Then
            grdError.AddItem ""
            grdError.TextMatrix(grdError.Rows - 1, 0) = i + 1
            grdError.TextMatrix(grdError.Rows - 1, 1) = sitem
            grdError.TextMatrix(grdError.Rows - 1, 2) = "ﬂÊœ €Ì— „ÊÃÊœ"
            If Not grdError.Visible Then grdError.Visible = True
        ElseIf grdError.Rows = 1 Then
            If nAdd >= 1000 Then
                cString.Shorten 1
                strTab.Append cString.GetAsString
                cString.Clear
                nAdd = 0
            End If
        
            nAdd = nAdd + 1
            If cString.length = 0 Then
                cString.Append "insert into FILE6_20(" & _
                    "DOC_NO," & _
                    "ITEM," & _
                    "QUANT," & _
                    "PRICE_C," & _
                    "PRICE," & _
                    "[MAN]," & _
                    "[USER_IP]," & _
                    "[COST]" & _
                    ")" & _
                    "VALUES"
            End If
            cString.Append "("
            cString.Append addstring(xDoc_No.text) & ","
            cString.Append cm.Parameters("@ITEM").Value & ","
            cString.Append nquant & ","
            cString.Append cm.Parameters("@PRICE").Value & ","
            cString.Append nPrice & ","
            cString.Append addstring("0001") & ","
            cString.Append addstring(GetComputerNamecIpName) & ","
            cString.Append cm.Parameters("@COST").Value
            cString.Append ")"
            cString.Append ","
        End If
    End If
Next

If grdError.Rows > 1 Then
    GoTo Finally
End If


If cString.length > 0 Then
    cString.Shorten 1
    strTab.Append cString.GetAsString()
End If

nCount = 1
con.BeginTrans
On Error GoTo myerror

For i = 0 To strTab.Count - 1 Step nCount
    Me.Caption = sCaption & " - " & IIf(i + nCount > strTab.Count, strTab.Count, i + nCount) & " „‰ " & strTab.Count
    prog1.Value = Round((i + 1) / strTab.Count, 2) * 100
    con.Execute strTab.GetStrings(i, nCount, 1), nAffect
Next

'UpdateInvTotal xDoc_No.text, con
con.CommitTrans
getCSV = True
Finally:
prog1.Visible = False
Me.Caption = sCaption
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
If TransCount(con) > 0 Then con.RollbackTrans
GoTo Finally
End Function
Private Function getCSVOld()
Dim cFileName As String
Dim sb As New ChilkatStringBuilder

Common1.InitDir = ""
Common1.FileName = ""
Common1.Filter = "CSV (*.CSV*)|*.CSV*"
Common1.ShowOpen

cFileName = Common1.FileName

If cFileName = "" Then Exit Function

Dim cSv As New ChilkatCsv
cSv.Delimiter = ","

cSv.HasColumnNames = 1
nAccess = sb.LoadFile(cFileName, "UTF-8")
If nAccess = 0 Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰  Õ„Ì· «·„·›"
    Exit Function
End If

nAccess = cSv.LoadFromString(sb.GetAsString)

If cSv.NumRows < 1 Then Exit Function

Dim Tb As New ChilkatStringBuilder
Dim cString As New ChilkatStringBuilder
Dim loctable As New ADODB.RecordSet

prog1.Visible = True

Dim sCaption As String
Dim bSkip As Boolean
Dim sitem As String
Dim nquant As Long
Dim nPrice As Double
Dim i As Long


sCaption = Me.Caption
grdError.Rows = 1
con.BeginTrans
For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
    
    sitem = Trim(cSv.GetCell(i, 0))
    nquant = Val(cSv.GetCell(i, 1))
    nPrice = Val(cSv.GetCell(i, 2))
    
    bSkip = False
    If Trim(sitem) = "" Or Len(sitem) < 3 Or Len(sitem) > 30 Then
        bSkip = True
            grdError.AddItem ""
            grdError.TextMatrix(grdError.Rows - 1, 0) = i
            grdError.TextMatrix(grdError.Rows - 1, 1) = sitem
            grdError.TextMatrix(grdError.Rows - 1, 2) = "ﬂÊœ €Ì— „ÊÃÊœ"
            bSkip = True
    End If
        
    If Not bSkip Then
        Dim cm As New ADODB.command
        Set cm = cmd("dbo.sp_item", con, adStoredProc, AddFlag(Empty, "BARCODE_FIND", sitem))
        cm.Execute
        
        If IsNull(cm.Parameters("@ITEM").Value) Then
            grdError.AddItem ""
            grdError.TextMatrix(grdError.Rows - 1, 0) = i
            grdError.TextMatrix(grdError.Rows - 1, 1) = sitem
            grdError.TextMatrix(grdError.Rows - 1, 2) = "ﬂÊœ €Ì— „ÊÃÊœ"
            bSkip = True
        ElseIf nquant = 0 Then
            grdError.AddItem ""
            grdError.TextMatrix(grdError.Rows - 1, 0) = i
            grdError.TextMatrix(grdError.Rows - 1, 1) = sitem
            grdError.TextMatrix(grdError.Rows - 1, 2) = "·«  ÊÃœ ﬂ„Ì…"
            bSkip = True
        End If
        
        If Not bSkip Then
            If nAdd >= 1000 Then
                cString.Shorten 1
                con.Execute cString.GetAsString
                cString.Clear
                nAdd = 0
            End If
        
            nAdd = nAdd + 1
            If cString.length = 0 Then
                cString.Append "insert into FILE6_20(" & _
                    "DOC_NO," & _
                    "ITEM," & _
                    "QUANT," & _
                    "PRICE," & _
                    "[MAN]," & _
                    "[USER_IP]," & _
                    "[COST]" & _
                    ")" & _
                    "VALUES"
            End If
            cString.Append "("
            cString.Append addstring(xDoc_No.text) & ","
            cString.Append cm.Parameters("@ITEM").Value & ","
            cString.Append nquant & ","
            cString.Append nPrice & ","
            cString.Append addstring("0001") & ","
            cString.Append addstring(GetComputerNamecIpName) & ","
            cString.Append cm.Parameters("@ITEM").Value
            cString.Append ")"
            cString.Append ","
        End If
    
    End If
Next
If cString.length > 0 Then
    cString.Shorten 1
    con.Execute cString.GetAsString()
    'UpdateInvTotal xDoc_No.text, con
End If
con.CommitTrans
Me.Caption = sCaption
getCSV = True
Finaly:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
If TransCount(con) > 0 Then con.RollbackTrans
GoTo Finaly
End Function
Private Sub myloadgrd2()
Dim cString As String
cString = "SELECT FILE1_10SC.CODE ," & _
           "FILE1_10SC.DESCA ," & _
           " FILE1_10SC.BARCODE_GS1_SEC," & _
           " SUM(FILE6_20.QUANT)," & _
           " ROUND(SUM(FILE6_20.TOTAL)/SUM(FILE6_20.QUANT) ,2)," & _
           " SUM(FILE6_20.TOTAL ) " & _
           " FROM FILE6_20 " & _
           " INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO " & _
           " INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM" & _
           " INNER JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE" & _
           " WHERE FILE6_20.QUANT <> 0 "
If xIs_collect.Value = 1 Then
    cString = cString & " AND INV_NO = " & MyParn(xDoc_No.text)
Else
    cString = cString & " AND DOC_NO = " & MyParn(xDoc_No.text)
End If
cString = cString & " GROUP BY FILE1_10SC.CODE," & _
           " FILE1_10SC.DESCA," & _
           " FILE1_10SC.BARCODE_GS1_SEC" & _
           " ORDER BY FILE1_10SC.CODE "
Set data12.RecordSet = cmd(cString, con).Execute
Fixgrd2
End Sub
Private Sub Fixgrd2()
With grid2
.FixedCols = 0
.Cols = 6
'                0        1         2           3               4           5       6
.FormatString = "ﬂÊœ|" & "«·ﬁ”„|" & "BARCODE GS1|" & "«·⁄œœ|" & "«·”⁄—|" & "«·≈Ã„«·Ï"
.RowHeight(0) = 800
.WordWrap = True
.ColWidth(0) = 1000
.ColWidth(1) = 3000
.ColWidth(2) = 2000
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 1500
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
'.Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter
.ShowCell .Rows - 1, 0
End With
End Sub
Private Sub cmdSign_Click()

If Not skipPortal Then
    Exit Sub
End If

AddLod_Data cusername, 0, " «—”«· ›« Ê—…  ›’Ì·Ì… " & Me.Caption, con, xDoc_No.text, xDate.text, , xCodeDesca.Caption

Me.MousePointer = vbHourglass
cmdSign.Enabled = False
cmdSignGroup.Enabled = False
SendInvoice
Handlecontrols xDoc_No.Tag
Me.MousePointer = vbNormal
End Sub
Private Sub cmdSignGroup_Click()
If Not skipPortal Then
    Exit Sub
End If

AddLod_Data cusername, 0, " «—”«· ›« Ê—… ≈Ã„«·Ì… " & Me.Caption, con, xDoc_No.text, xDate.text, , xCodeDesca.Caption

Me.MousePointer = vbHourglass
cmdSign.Enabled = False
cmdSignGroup.Enabled = False
SendInvoice True
Me.MousePointer = vbNormal
End Sub
Private Function SendInvoice(Optional bGroup As Boolean = False) As Boolean
Dim cError As String
bsuccess = SubmitDocument(xDoc_No.text, con, cError, Me, IIf(lIsRet, "C", "I"), bGroup)

If bsuccess Then
    eInform " „ «—”«· «·›« Ê—… «·«·Ìﬂ —Ê‰Ì… »‰Ã«Õ"
    If Not openCardTable(tbMode.tbFind, xDoc_No.text) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
ElseIf cError <> "" Then
    MsgBox cError
Else
    eInform "Œÿ« «À‰«¡ «—”«· «·›« Ê—…"
End If
End Function
Private Function addNewDoc() As String
If optType(1).Value Or optType(2).Value Or optType(3).Value Then
    addNewDoc = IncData("FILE6_20H", "DOC_NO", con, "FILE6_20H.INV_TYPE = 2")
    If addNewDoc = "" Then addNewDoc = RetZero("1", 12)
ElseIf optType(5).Value Then
    addNewDoc = IncData("FILE6_20H", "DOC_NO", con, "FILE6_20H.INV_TYPE = 3")
    If addNewDoc = "" Then addNewDoc = "1" & RetZero("1", 11)
ElseIf optType(6).Value Then
    addNewDoc = IncData("FILE6_20H", "DOC_NO", con, "FILE6_20H.INV_TYPE = 4")
    If addNewDoc = "" Then addNewDoc = "2" & RetZero("1", 11)
End If
End Function
Private Function ValidQuant() As Boolean
Dim nType As Integer
For i = 1 To grid1.Rows - 1
    If grid1.ValueMatrix(i, 10) = 0 Then
    ElseIf nType = 0 Then
        nType = IIf(grid1.ValueMatrix(i, 10) > 0, 1, 2)
    Else
        If nType <> IIf(grid1.ValueMatrix(i, 10) > 0, 1, 2) Then Exit Function
    End If
Next
ValidQuant = True
End Function
Private Function TransToTaxInv() As Boolean
Dim loctable As New ADODB.RecordSet

Dim cString As String
cString = "SELECT ID" & _
          " FROM (SELECT ID, item, SUM(quant) as NetQty" & _
          " FROM FILE6_20 " & _
          "  GROUP BY ID, item" & _
          ") AS ItemSums" & _
          " WHERE DOC_NO = " & MyParn(xDoc_No.text) & _
          " GROUP BY ID" & _
          " HAVING MAX(NetQty) > 0 AND MIN(NetQty) < 0"

Set loctable = cmd(cString, con).Execute
If Not loctable.EOF Then
    MsgBox "«·›« Ê—…  Õ ÊÌ ⁄·Ì ﬂ„Ì… ”«·Ì… Ê„ÊÃ»…"
    Exit Function
End If

loctable.Close
cString = "SELECT TOP 1 FILE6_20H.*" & _
          " FROM FILE6_20 " & _
          " INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO " & _
          " WHERE FILE6_20.DOC_NO = " & MyParn(xDoc_No.text)

Set loctable = cmd(cString, con).Execute
If loctable.EOF Then
    MsgBox "·«  ÊÃœ ”Ã·«  · ÕÊÌ·Â«"
    Exit Function
End If

If loctable!INV_TYPE & "" <> 2 Then
    MsgBox "«·„” ‰œ«  ·Ì” «–‰"
    Exit Function
End If

Dim aInsert As Variant
Dim sNewDoc As String
If loctable!TOTAL_ITEM >= 0 Then
    sNewDoc = IncData("FILE6_20H", "DOC_NO", con, "FILE6_20H.INV_TYPE = 3")
    If sNewDoc = "" Then sNewDoc = "1" & RetZero("1", 11)
Else
    sNewDoc = IncData("FILE6_20H", "DOC_NO", con, "FILE6_20H.INV_TYPE = 4")
    If sNewDoc = "" Then sNewDoc = "2" & RetZero("1", 11)
End If

aInsert = AddFlag(Empty, "DOC_NO", addstring(sNewDoc))
aInsert = AddFlag(aInsert, "DATE", addDate(myFormat(Date)))
aInsert = AddFlag(aInsert, "CODE", addstring(loctable!code))
aInsert = AddFlag(aInsert, "PO_NO", addstring(loctable!po_no))
aInsert = AddFlag(aInsert, "USERNAME", addstring(cusername))
aInsert = AddFlag(aInsert, "BRANCH", addstring("00"))
aInsert = AddFlag(aInsert, "ISINVOICE", "1")
aInsert = AddFlag(aInsert, "STORE", addstring(loctable!STORE))
aInsert = AddFlag(aInsert, "NOTES", addstring(loctable!doc_no))

con.BeginTrans
con.Execute addInsert(aInsert, "FILE6_20H")


con.Execute "INSERT INTO FILE6_20(" & _
            "DOC_NO," & _
            "ITEM," & _
            "QUANT," & _
            "PRICE_C," & _
            "PRICE," & _
            "MAN," & _
            "COST)" & _
            " SELECT " & MyParn(sNewDoc) & "," & _
            " FILE6_20.ITEM," & _
            " FILE6_20.QUANT," & _
            " PRICE_C," & _
            " ROUND((PRICE/1.14) * ( 1 - FILE6_20H.DISCOUNT_RATE),2) ," & _
            "'0001'," & _
            " FILE6_20.COST" & _
            " FROM FILE6_20 " & _
            " INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO" & _
            " WHERE FILE6_20.DOC_NO = " & MyParn(loctable!doc_no)
            
con.Execute "UPDATE FILE6_20H SET " & _
            " FILE6_20H.TAX = ROUND(FILE6_20H.TOTAL_ITEM * 0.14,2)" & _
            " FROM FILE6_20H WHERE FILE6_20H.DOC_NO = " & MyParn(sNewDoc)

con.Execute "UPDATE FILE6_20H SET " & _
            " FILE6_20H.DOC_NO_TAX = " & addstring(sNewDoc) & "," & _
            " FILE6_20H.DISCOUNT = 0," & _
            " FILE6_20H.TAX = 0" & _
            " FROM FILE6_20H WHERE FILE6_20H.DOC_NO = " & MyParn(loctable!doc_no)

con.Execute "DELETE FROM FILE6_20 WHERE DOC_NO = " & MyParn(loctable!doc_no)
con.CommitTrans
TransToTaxInv = True
If loctable!TOTAL_ITEM >= 0 Then
    optType(5).Value = True
Else
    optType(6).Value = True
End If
If Not openCardTable(tbMode.tbFind, sNewDoc) Then
    If Not openCardTable Then
        myDefine
    End If
End If
End Function
Private Function TransToTaxInvOld()
Dim loctable As New ADODB.RecordSet

Dim cString As String
cString = "SELECT ID" & _
          " FROM (SELECT ID, item, SUM(quant) as NetQty" & _
          " FROM FILE6_20 " & _
          "  GROUP BY ID, item" & _
          ") AS ItemSums" & _
          " WHERE DOC_NO = " & MyParn(xDoc_No.text) & _
          " GROUP BY ID" & _
          " HAVING MAX(NetQty) > 0 AND MIN(NetQty) < 0"

Set loctable = cmd(cString, con).Execute
If Not loctable.EOF Then
    MsgBox "«·›« Ê—…  Õ ÊÌ ⁄·Ì ﬂ„Ì… ”«·Ì… Ê„ÊÃ»…"
    Exit Function
End If

loctable.Close
cString = "SELECT TOP 1 FILE6_20H.*" & _
          " FROM FILE6_20 " & _
          " INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO " & _
          " WHERE FILE6_20.DOC_NO = " & MyParn(xDoc_No.text)

Set loctable = cmd(cString, con).Execute
If loctable.EOF Then
    MsgBox "·«  ÊÃœ ”Ã·«  · ÕÊÌ·Â«"
    Exit Function
End If

If loctable!INV_TYPE & "" <> 2 Then
    MsgBox "«·„” ‰œ«  ·Ì” «–‰"
    Exit Function
End If

Dim aInsert As Variant
Dim sNewDoc As String
If loctable!TOTAL_ITEM >= 0 Then
    sNewDoc = IncData("FILE6_20H", "DOC_NO", con, "FILE6_20H.INV_TYPE = 3")
    If sNewDoc = "" Then sNewDoc = "1" & RetZero("1", 11)
Else
    sNewDoc = IncData("FILE6_20H", "DOC_NO", con, "FILE6_20H.INV_TYPE = 4")
    If sNewDoc = "" Then sNewDoc = "2" & RetZero("1", 11)
End If
aInsert = AddFlag(Empty, "DOC_NO", addstring(sNewDoc))
aInsert = AddFlag(aInsert, "DATE", addDate(myFormat(Date)))
aInsert = AddFlag(aInsert, "CODE", addstring(loctable!code))
aInsert = AddFlag(aInsert, "PO_NO", addstring(loctable!po_no))
aInsert = AddFlag(aInsert, "BRANCH", addstring("00"))
aInsert = AddFlag(aInsert, "ISINVOICE", "1")
aInsert = AddFlag(aInsert, "STORE", addstring(loctable!STORE))
aInsert = AddFlag(aInsert, "NOTES", addstring(loctable!doc_no))

con.BeginTrans
con.Execute addInsert(aInsert, "FILE6_20H")

con.Execute "INSERT INTO FILE6_20(" & _
            "DOC_NO," & _
            "ITEM," & _
            "QUANT," & _
            "PRICE_C," & _
            "PRICE," & _
            "MAN," & _
            "COST)" & _
            " SELECT " & MyParn(sNewDoc) & "," & _
            " FILE6_20.ITEM," & _
            " FILE6_20.QUANT," & _
            " FILE6_20.PRICE_C," & _
            " ROUND((FILE6_20.PRICE / 1.14) * ( 1 - FILE6_20H.DISCOUNT_RATE),6)," & _
            "'0001'," & _
            " COST" & _
            " FROM FILE6_20 " & _
            " INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO" & _
            " WHERE FILE6_20.DOC_NO = " & MyParn(loctable!doc_no)

con.Execute "UPDATE FILE6_20H SET " & _
            " FILE6_20H.TAX = ROUND(FILE6_20H.TOTAL_ITEM * 0.14,2)" & _
            " FROM FILE6_20H WHERE FILE6_20H.DOC_NO = " & MyParn(sNewDoc)

con.Execute "UPDATE FILE6_20H SET " & _
            " FILE6_20H.DOC_NO_TAX = " & addstring(sNewDoc) & "," & _
            " FILE6_20H.DISCOUNT = 0," & _
            " FILE6_20H.TAX = 0" & _
            " FROM FILE6_20H WHERE FILE6_20H.DOC_NO = " & MyParn(loctable!doc_no)
con.Execute "DELETE FROM FILE6_20 WHERE DOC_NO = " & MyParn(loctable!doc_no)
con.CommitTrans

If loctable!TOTAL_ITEM >= 0 Then
    optType(5).Value = True
Else
    optType(6).Value = True
End If
If Not openCardTable(tbMode.tbFind, sNewDoc) Then
    If Not openCardTable Then
        myDefine
    End If
End If
End Function
Private Sub doprint_SUBINV()
Dim aHeader(2)
If Not myValid Then Exit Sub
Dim temptable As New ADODB.RecordSet
Dim sourcetable As New ADODB.RecordSet
Dim lPrice As Boolean
Dim nFBal As Double
Dim nPay As Double
Dim nBal As Double, nDisc As Double, nTotalItem As Double, nRate As Double, nTotalQuant As Double
Dim cAddres As String, cTaxNo As String

nDisc = Val(GetDesca("SELECT SUM(DISCOUNT) FROM FILE6_20H  WHERE ISINVOICE = 0 AND CODE = " & MyParn(xCode.text) & " AND DATE >= " & DateSq(XDATESALES1.text) & " AND DATE <= " & DateSq(XDATESALES2.text), con) & "") + Val(xDiscount.text)
nTotalItem = Val(GetDesca("SELECT SUM(FILE6_20.TOTAL) FROM FILE6_20H  INNER JOIN FILE6_20 ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO WHERE FILE6_20H.CODE = " & MyParn(xCode.text) & " AND FILE6_20H.DATE >= " & DateSq(XDATESALES1.text) & " AND FILE6_20H.DATE <= " & DateSq(XDATESALES2.text), con) & "")
nRate = 1 - (Round(nDisc / nTotalItem, 10))
nTotalQuant = Val(GetDesca("SELECT SUM(QUANT) FROM FILE6_20H  INNER JOIN FILE6_20 ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO WHERE FILE6_20H.CODE = " & MyParn(xCode.text) & " AND FILE6_20H.DATE >= " & DateSq(XDATESALES1.text) & " AND FILE6_20H.DATE <= " & DateSq(XDATESALES2.text), con) & "")

Dim cString As String
cString = "SELECT FILE6_20.ITEM ," & _
          "FILE1_10.DESCA ," & _
          "FILE1_10.SCAL ," & _
          "FILE1_10.C_SCAL," & _
          "FILE1_10.COLOR ," & _
          "FILE3_10.MANAGER," & _
          "SUM(FILE6_20.QUANT ) AS QUANT ," & _
          "FILE1_10.MODELFACT0," & _
          "SUM(FILE6_20.TOTAL) AS TOTAL " & _
          " FROM FILE1_10 " & _
          " INNER JOIN FILE6_20 ON FILE6_20.ITEM = FILE1_10.ITEM" & _
          " INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO" & _
          " INNER JOIN FILE3_10 ON FILE6_20H.CODE = FILE3_10.CODE" & _
          " WHERE FILE6_20H.CODE = " & MyParn(xCode.text) & _
          " AND FILE6_20H.DATE >= " & DateSq(XDATESALES1.text) & _
          " AND FILE6_20H.DATE <= " & DateSq(XDATESALES2.text) & _
          " GROUP BY FILE6_20.ITEM ," & _
          " FILE1_10.DESCA," & _
          " FILE1_10.SCAL," & _
          " FILE1_10.C_SCAL," & _
          " FILE1_10.COLOR," & _
          " FILE1_10.MODELFACT0," & _
          " FILE3_10.MANAGER"

Set sourcetable = cmd(cString, con).Execute

cAddres = GetDesca("SELECT ADDRESS FROM FILE3_10 WHERE CODE = " & MyParn(xCode.text), con)
cTaxNo = GetDesca("SELECT EMAIL FROM FILE3_10 WHERE CODE = " & MyParn(xCode.text), con)

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

With sourcetable
Do While Not .EOF
    If !Quant <> 0 Then
        temptable.AddNew
        temptable!str1 = xDoc_no2.Caption
        temptable!str2 = IIf(sourcetable!Manager & "" = "", xCodeDesca.Caption, sourcetable!Manager)
        temptable!Date1 = xDate.text
        temptable!STR6 = TurnValue(cAddres)
        temptable!str15 = TurnValue(cTaxNo)
        temptable!str14 = !MODELFACT0
        temptable!VAL2 = Round((((!TOTAL / !Quant) * nRate) / 1.14), 6)
        temptable!VAL3 = Round((((!TOTAL / !Quant) * nRate) / 1.14), 6) * !Quant
        temptable!VAL1 = !Quant
        temptable!str8 = !DESCA
        temptable!str9 = !SCAL
        temptable!VAL19 = !C_SCAL
        temptable!STR7 = !color
        temptable!str3 = " ›« Ê—… „»Ì⁄«  "
        
    '   temptable!VAL7 = Val(xTotalItem.Caption)
        temptable!val9 = nTotalQuant
        
    '   temptable!str18 = "≈Ã„«·Ï «’‰«›"
    '   temptable!val18 = Val(xTotalItem.Caption)
    '   temptable!val17 = 0
    '   temptable!Val8 = Val(xTotDisc.Caption)
        
        temptable!val15 = (nTotalItem * nRate) - Round((nTotalItem * nRate) / 1.14, 2)
        temptable!val16 = nTotalItem * nRate
    '   temptable!str16 = MyOnly(Val(xTotal.Text))
    '   temptable!str18 = "Œ’„ «·›« Ê—…"
    '   temptable!VAL7 = Val(xDiscount.Text)
        
        temptable.Update
    End If
    .MoveNext
Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
REPORT1.ReportFileName = App.Path & "\Reports\PRINT_INVOICE.rpt"
REPORT1.Destination = crptToWindow
REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Function skipPortal() As Boolean
Dim cmSkip As ADODB.command
aPrm = AddFlag(aPrm, "doc_no", xDoc_No.text)
Set cmSkip = cmd("[dbo].[sp_skip_portal]", con, adStoredProc, aPrm)
cmSkip.Execute
If Not IsNull(cmSkip.Parameters("@DOC_NO_SKIP")) Then
    MsgBox "›« Ê—… —ﬁ„ " & cmSkip.Parameters("@DOC_NO_SKIP") & " ·„  —”· ··»Ê— «·"
    Exit Function
End If
skipPortal = True
End Function
Public Sub myProcOrder(pDoc_no As String)
If Not openCardTable(tbMode.tbFind, pDoc_no) Then
    If Not openCardTable Then
        myDefine
    End If
End If
End Sub

Public Function fnBalance(pItem As String, con As ADODB.Connection, Optional pstore As String = "", Optional pDate As String = "", Optional pId As String = "") As Long
Dim cmBalance As New ADODB.command
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "ITEM", pItem)
If pstore <> "" Then aPrm = AddFlag(aPrm, "STORE", pstore)
If IsDate(pDate) Then aPrm = AddFlag(aPrm, "DATE", myFormat_sp(pDate))
If pId <> "" Then aPrm = AddFlag(aPrm, "ID", pId)

Set cmBalance = cmd("dbo.sp_balance", con, adStoredProc, aPrm)
cmBalance.Execute

fnBalance = Val(cmBalance.Parameters("@BALANCE") & "")
Set cmBalance = Nothing
End Function
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
cString = "SELECT TOP 1 FILE6_20H.*," & _
          " FILE3_10.DESCA AS CLIENTDESCA, " & _
          " FILE6_53H.NOTES AS LIST_DESCA " & _
          " FROM FILE6_20H " & _
          " INNER JOIN FILE3_10 ON FILE3_10.CODE = FILE6_20H.code" & _
          " INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE" & _
          " LEFT JOIN FILE6_53H ON FILE6_20H.PRICE_LIST = FILE6_53H.DOC_NO"

  
If pMode = tbMode.tbFirst Then
    'cOrder = "Order by FILE6_20H.DOC_NO_FLAG"
    cOrder = "Order by FILE6_20H.DOC_NO"
ElseIf pMode = tbMode.tblast Then
    'cOrder = "Order by FILE6_20H.DOC_NO_FLAG DESC"
    cOrder = "Order by FILE6_20H.DOC_NO DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "FILE6_20H.Doc_no = " & MyParn(pDoc_no)
ElseIf pMode = tbMode.tbPrevious Then
    'cWhere = "FILE6_20H.DOC_NO_FLAG < " & MyParn(pDoc_no)
    'cOrder = "Order by FILE6_20H.DOC_NO_FLAG DESC"
    
    cWhere = "FILE6_20H.DOC_NO < " & MyParn(pDoc_no)
    cOrder = "Order by FILE6_20H.DOC_NO DESC"
ElseIf pMode = tbMode.tbNext Then
    'cWhere = "FILE6_20H.DOC_NO_FLAG > " & MyParn(pDoc_no)
    cWhere = "FILE6_20H.DOC_NO > " & MyParn(pDoc_no)
    cOrder = "Order by FILE6_20H.DOC_NO"
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
Private Function ChangePriceList() As Boolean
Dim db As New clsDb
Dim rsPrice As New ADODB.RecordSet
Set rsPrice = db.myRs("select * from file6_53 where doc_no = " & xPrice_list.Tag): If rsPrice Is Nothing Then GoTo CleanUp

sCaption = Me.Caption
prog1.Visible = True

Dim nPrice As Variant
For i = 1 To grid1.Rows - 2
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & grid1.Rows - 2
    prog1.Value = Round(i / (grid1.Rows - 2), 2) * 100
    
    rsPrice.Find "item = " & MyParn(grid1.TextMatrix(i, 1)), , adSearchForward, adBookmarkFirst
    If Not rsPrice.EOF Then
         db.Sql = "update file6_20 " & _
                    " set file6_20.price = " & rsPrice!price & _
                    " where file6_20.doc_no = " & xDoc_No.text & _
                    " and file6_20.item = " & MyParn(grid1.TextMatrix(i, 1))
    End If
Next
db.Sql = "UPDATE FILE6_20H SET PRICE_LIST = " & xPrice_list.Tag & " WHERE DOC_NO = " & MyParn(xDoc_No.text)
If Not db.ExecuteTransaction Then GoTo CleanUp
ChangePriceList = True
CleanUp:
Set db = Nothing
prog1.Visible = False
Me.Caption = sCaption
End Function
