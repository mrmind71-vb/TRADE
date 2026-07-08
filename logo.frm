VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form logofrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "‘⁄«— «·‘—ﬂ…"
   ClientHeight    =   6390
   ClientLeft      =   420
   ClientTop       =   405
   ClientWidth     =   8175
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
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
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   8175
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   90
      Width           =   7890
      Begin Threed.SSCommand cmdSave 
         Height          =   510
         Left            =   3960
         TabIndex        =   15
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "logo.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "logo.frx":29F5
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "logo.frx":528E
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   1350
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "logo.frx":75B1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "logo.frx":9D4D
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   510
         Left            =   2655
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "logo.frx":C1E1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "logo.frx":E422
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   510
         Left            =   5265
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "logo.frx":1070F
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "logo.frx":12717
      End
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   6570
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "logo.frx":146CE
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "logo.frx":16A99
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   5310
      Width           =   3480
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2610
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   180
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
         Picture         =   "logo.frx":18B42
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "logo.frx":1ACE9
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1710
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   180
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
         Picture         =   "logo.frx":1CD30
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "logo.frx":1EE1B
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   855
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   180
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
         Picture         =   "logo.frx":20E15
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "logo.frx":22F26
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   180
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
         Picture         =   "logo.frx":24F20
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "logo.frx":27144
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   3480
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1845
      Width           =   7845
      Begin Threed.SSCommand cmdScan 
         Height          =   555
         Left            =   3285
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2835
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   979
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
         Picture         =   "logo.frx":29215
         Caption         =   "„”Õ ÷Ê∆Ì"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "logo.frx":2B71D
      End
      Begin Threed.SSCommand cmdDelPhoto 
         Height          =   555
         Left            =   90
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2835
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   979
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
         Picture         =   "logo.frx":2E2B8
         Caption         =   "Õ–› «·’Ê—…"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "logo.frx":3087D
      End
      Begin Threed.SSCommand cmdPhoto 
         Height          =   555
         Left            =   1710
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2835
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   979
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
         Picture         =   "logo.frx":3334E
         Caption         =   "„‰ „·›"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "logo.frx":35593
      End
      Begin ImgXCtrl6.ImgXCtrl ImgX1 
         Height          =   2580
         Left            =   90
         TabIndex        =   25
         Top             =   180
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   4551
         BackColor       =   16777215
         BorderStyle     =   4
         AutoZoom        =   -1  'True
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "mrmind"
         LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   810
      Width           =   7845
      Begin VB.TextBox xDescA 
         Alignment       =   1  'Right Justify
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
         Left            =   2025
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   4695
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   5400
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "2"
         Top             =   180
         Width           =   1320
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ"
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
         Left            =   6795
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "»Ì«‰ «·‘⁄«—"
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
         Left            =   6795
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   585
         Width           =   840
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   855
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
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
      Top             =   225
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
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
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
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
   Begin MSAdodcLib.Adodc DATA5 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
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
      Width           =   1590
      _ExtentX        =   2805
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
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   6015
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   661
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   22
         Top             =   45
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   476
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
         Left            =   3465
         TabIndex        =   23
         Top             =   45
         Width           =   3465
         _ExtentX        =   6112
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
         Left            =   6975
         TabIndex        =   24
         Top             =   45
         Width           =   3465
         _ExtentX        =   6112
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
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
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
End
Attribute VB_Name = "logofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim cFilter As String
Dim formMode As Byte
Dim oSearch As New Search_abd
Dim CardTable As ADODB.Recordset
Dim WithEvents twain As ImgXTwain
Attribute twain.VB_VarHelpID = -1
Const LoadMode = 1, DefineMode = 2

Private Sub cmdDelPhoto_Click()
On Error GoTo myerror
con.Execute "UPDATE LOGO_CODES SET IMAGE = NULL WHERE CODE = " & xCode.text
ImgX1.Images.Clear
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub cmdPhoto_Click()
    Dim cFile As String, cNewFile As String
    Common1.FileName = ""
    'Common1.InitDir = doc_dir()
    Common1.Filter = "Pictures (*.Jpg)|*.Jpg"
    Common1.ShowOpen
    If Common1.FileTitle <> "" Then
        cFile = Common1.FileName
        If cFile <> "" Then
            myReplaceImage xCode.text, cFile
        End If
    End If
End Sub

Private Sub cmdScan_Click()
ScanImage
End Sub
Private Sub ScanImage()
On Error GoTo myerror
Set twain = New ImgXTwain
twain.OpenTwain Me.hwnd
If twain.QuerySupport(ixtcResolution) Then
     twain.Resolution = 150
End If
If twain.Sources.Count > 1 Then twain.SelectSource
twain.Acquire False, Me.hwnd
Exit Sub
myerror:
MsgBox Err.Number & vbCrLf & Err.Description
Err.Clear
End Sub

Private Sub Twain_ImageAcquired(Image As ImgX_Image)
ImgX1.Images.Replace Image, , False
ImgX1.TIFCompression = ixtcJPEGCompression

myReplaceImage xCode.text
End Sub
Private Sub Twain_TwainError(ByVal erNum As Long, ByVal erSource As String, ByVal Description As String)
MsgBox "Error Number:  " & erNum & vbCrLf & Description, vbInformation, erSource
End Sub
Private Sub Twain_CanCloseTwain()
    ' This event is called after you call Acquire.
    ' It let's you know when it's safe to call CloseTwain.
    twain.CloseTwain
    ' Steps menu
End Sub
Private Function myReplaceImage(sCode As String, Optional pFile As String = "") As Boolean
If pFile <> "" Then
    myReplaceImage = SaveImageToDB(pFile, ImgX1, "LOGO_CODES", "IMAGE", "CODE", sCode, con)
Else
    myReplaceImage = SaveImageToDB("", ImgX1, "LOGO_CODES", "IMAGE", "CODE", sCode, con)
End If

If Not openCardTable(tbMode.tbFind, xCode.text) Then
    If Not openCardTable Then
        myDefine
    End If
End If
End Function
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End If
End Sub
Private Sub Form_Load()
openCon con
        
If Not openCardTable Then myDefine
End Sub
Private Sub CmdAdd_Click()
myDefine
xDescA.SetFocus
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    con.Execute "Delete  From logo_codes  Where code = " & MyParn(xCode.text)
    If Not openCardTable(tbMode.tbPrevious, xCode.text) Then
        If Not openCardTable(tbMode.tbFirst) Then
            myDefine
        End If
    End If
End If
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Not myValid Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ »Ì«‰«  «·„Ê—œ »‰Ã«Õ"
If Not openCardTable(tbMode.tbFind, xCode.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub CmdUndo_Click()
myUndo
End Sub
Private Sub CmdInform_Click()
CardLookup
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
Private Sub CmdNext_Click()
If Not openCardTable(tbMode.tbNext, xCode.text) Then
    If Not openCardTable(tbMode.tblast) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdPrevious_Click()
If Not openCardTable(tbMode.tbPrevious, xCode.text) Then
       If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
End Sub
Sub Handlecontrols(nMode)
cmdAdd.Enabled = (nMode = LoadMode)
cmddel.Enabled = (nMode = LoadMode)
cmdInform.Enabled = (nMode = LoadMode)
cmdScan.Enabled = nMode = LoadMode
cmdPhoto.Enabled = nMode = LoadMode
cmdDelPhoto.Enabled = nMode = LoadMode

xCode.Tag = nMode
Dim nRecord As Long, nRecords As Long
retRecords xCode.text, nRecords, nRecord

If nMode = LoadMode Then
    panel1(0).Caption = ArbString("”Ã· " & nRecord & " „‰ " & nRecords)
Else
    panel1(0).Caption = "«÷«›… ”Ã· " & (nRecords + 1)
End If

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sDoc_no = ""
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From logo_codes"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ «·‘⁄«—« "
oSearch.Show 1
End Sub
Sub myDefine()
Dim sCodeNew As String

xCode.text = Newflag("LOGO_CODES", "CODE", con) & ""
xDescA.text = ""
ImgX1.Images.Clear

Handlecontrols DefineMode
End Sub
Sub myload()
xCode.text = CardTable!code & ""
xDescA.text = CardTable!desca
If IsNull(CardTable!Image) Then
    ImgX1.Images.Clear
Else
    ImgX1.Import.FromMemoryFile (CardTable!Image)
End If
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "DESCA", addstring(xDescA.text))
On Error GoTo myerror
If xCode.Tag = DefineMode Then
    Dim sCodeNew As String
    xCode.text = Newflag("LOGO_CODES", "CODE", con) & ""
    aInsert = AddFlag(aInsert, "CODE", addstring(xCode.text))
    con.Execute addInsert(aInsert, "LOGO_CODES")
Else
    con.Execute addUpdate(aInsert, "LOGO_CODES", "code = " & addstring(xCode.text))
End If
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Sub myProc()
If ActiveControl.Name = cmdInform.Name Then
    openCardTable tbMode.tbFind, oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    Unload oSearch
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set CardTable = Nothing
closeCon con
Set logofrm = Nothing
Err.Clear
End Sub

Private Sub xcode_LostFocus()
myLostFocus xCode
If Trim(xCode.text) = "" Or Not IsDgt(xCode.text) Then
     If xCode.Tag = LoadMode Then
        myDefine
    End If
Else
    If Not openCardTable(tbMode.tbFind, xCode.text) Then
       If xCode.Tag = LoadMode Then
            myDefine
       End If
    End If
End If
End Sub
Function myValid() As Boolean
If Not IsDgt(xCode.text) Then
    MsgBox "«·ﬂÊœ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If


If xDescA.text = "" Then
    MsgBox "«·≈”„ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

myValid = True
End Function
Private Sub myUndo()
If xCode.Tag = DefineMode Then
    If Not openCardTable Then
        myDefine
    End If
Else
    If Not openCardTable(tbMode.tbFind, xCode.text) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
End If
End Sub
Private Function openCardTable(Optional pMode As Integer = tbMode.tblast, Optional pCode As String = "", Optional bDefine As Boolean = True) As Boolean
Dim cString As String
Dim cWhere As String
Dim cOrder As String
If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pCode) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass

On Error GoTo myerror

If sCode <> "" Then
    cFilter = cFilter & " AND CODE = " & sCode
End If

cString = "SELECT TOP 1 * FROM LOGO_CODES"

If pMode = tbMode.tbFirst Then
    cOrder = "Order by CODE"
ElseIf pMode = tbMode.tblast Then
    cOrder = "Order by CODE DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "CODE = " & pCode
ElseIf pMode = tbMode.tbPrevious Then
    cWhere = "CODE < " & pCode
    cOrder = "Order by CODE desc"
ElseIf pMode = tbMode.tbNext Then
    cWhere = "CODE > " & pCode
    cOrder = "Order by CODE"
End If

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

If cFilter <> "" Then
    cString = cString & IIf(cWhere = "", " WHERE ", " AND ") & cFilter
End If

cString = cString & " " & cOrder

Set CardTable = mycmd(cString, con, adText)

If (Not CardTable.EOF) Then
    myload
    openCardTable = True
End If

Me.MousePointer = 0
Exit Function
myerror:
Me.MousePointer = vbNormal
MsgBox Err.Description
Err.Clear
End Function
Private Function retRecords(pCode, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If pCode <> "" Then
    cString = "SELECT Count(*) AS records,COUNT(CASE WHEN CODE <= " & pCode & " THEN 1 END) AS record"
Else
    cString = "SELECT Count(*) AS records,0 as record"
End If

cString = cString & " FROM logo_codes"
If cFilter <> "" Then cString = cString & " WHERE " & cFilter

Set loctable = mycmd(cString, con)
If Not loctable.EOF Then
    nRecords = loctable!RECORDS
    nRecord = Val(loctable!Record & "")
End If
End Function
Private Sub xDescA_GotFocus()
myGotFocus xDescA
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDescA
End Sub
Private Sub xcode_KeyPress(KeyAscii As Integer)
    ' Check if the key pressed is NOT a digit (0-9)
    ' and NOT the Backspace key (ASCII 8)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8) Then
        ' Setting KeyAscii to 0 cancels the input
        KeyAscii = 0
        'Beep  ' Optional: alert the user
    End If
End Sub
Private Sub xcode_GotFocus()
myGotFocus xCode
End Sub
