VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form CashOnLine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   12075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   12075
   ScaleWidth      =   5820
   Begin VB.Frame Frame2 
      Caption         =   "—›÷ «·≈” ·«„"
      Height          =   1770
      Left            =   990
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   8730
      Width           =   4605
      Begin VB.CheckBox XISRETONLINE 
         Alignment       =   1  'Right Justify
         Caption         =   "„— Ã⁄ «Ê‰ ·«Ì‰ „‰ ‘—ﬂ… ‘Õ‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   765
         Width           =   3030
      End
      Begin VB.TextBox XDATE_REC2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Left            =   90
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1170
         Width           =   1905
      End
      Begin VB.TextBox xnote_ret 
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
         Height          =   420
         Left            =   90
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   270
         Width           =   3030
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «” ·«„ „‰ «·‘Õ‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   420
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1260
         Width           =   2310
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "”»» «·—›÷"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   420
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   270
         Width           =   1860
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   180
      Top             =   -135
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
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
   Begin VB.Frame Frame1 
      Height          =   7035
      Left            =   990
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1710
      Width           =   4605
      Begin VB.CheckBox XNORET 
         Alignment       =   1  'Right Justify
         Caption         =   "»œÊ‰ —œ"
         Height          =   240
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1260
         Width           =   825
      End
      Begin VB.TextBox xAddress 
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
         Height          =   1140
         Left            =   135
         Locked          =   -1  'True
         MaxLength       =   1000
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   5805
         Width           =   4380
      End
      Begin VB.TextBox XNAME 
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
         Height          =   420
         Left            =   135
         Locked          =   -1  'True
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   5085
         Width           =   3210
      End
      Begin VB.TextBox XPHONE 
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
         Height          =   420
         Left            =   135
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   4605
         Width           =   2175
      End
      Begin VB.TextBox xDate_Pay 
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
         Height          =   420
         Left            =   135
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3666
         Width           =   2175
      End
      Begin VB.TextBox xShip_No 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   135
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   3153
         Width           =   2175
      End
      Begin VB.TextBox xcharge1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1080
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   1191
         Width           =   1230
      End
      Begin VB.TextBox xcharge2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   135
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1704
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo xSHIP 
         Height          =   390
         Left            =   135
         TabIndex        =   2
         Top             =   2700
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo XBANK 
         Height          =   390
         Left            =   135
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   4140
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "»‰ﬂ «·≈Ìœ«⁄ : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   4185
         Width           =   1275
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         Caption         =   "≈”„ «·⁄„Ì· : "
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
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   5130
         Width           =   1140
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         Caption         =   "⁄‰Ê«‰ «·⁄„Ì· "
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
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   5490
         Width           =   1050
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         Caption         =   " «—ÌŒ «·”œ«œ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   3735
         Width           =   1275
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         Caption         =   "—ﬁ„ »Ê·Ì’… : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   3240
         Width           =   1770
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         Caption         =   "—ﬁ„  ·Ì›Ê‰ «·⁄„Ì· :"
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
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   4635
         Width           =   1545
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         Caption         =   "‘—ﬂ… «·‘Õ‰ : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   2790
         Width           =   1545
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "≈Ã„«·Ï »Ê·Ì’… :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   2250
         Width           =   1590
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         Caption         =   "„’«—Ì› ‘Õ‰ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1260
         Width           =   1635
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "≈Ã„«·Ï ﬁÌ„… «·»Ê‰ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   765
         Width           =   1995
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "⁄œœ «·ﬁÿ⁄ «·„»«⁄… :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   315
         Width           =   1905
      End
      Begin VB.Label xtotalQuant 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   225
         Width           =   2175
      End
      Begin VB.Label xTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   708
         Width           =   2175
      End
      Begin VB.Label xTotalCharge 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   2217
         Width           =   2175
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "„’«—Ì› »Ê·Ì’… : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1755
         Width           =   1905
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
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
   Begin Threed.SSCommand CMD_SAVE 
      CausesValidation=   0   'False
      Height          =   510
      Left            =   3825
      TabIndex        =   8
      Top             =   10575
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   900
      _Version        =   196610
      CaptionStyle    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Õ›Ÿ «·»Ì«‰« "
      PictureAlignment=   3
   End
   Begin Threed.SSCommand CMD_EXIT 
      CausesValidation=   0   'False
      Height          =   510
      Left            =   1035
      TabIndex        =   24
      Top             =   10575
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   900
      _Version        =   196610
      CaptionStyle    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Œ—ÊÃ"
      PictureAlignment=   3
   End
   Begin Threed.SSCommand CMD_DELPAY 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3825
      TabIndex        =   32
      Top             =   11205
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   661
      _Version        =   196610
      CaptionStyle    =   1
      BackColor       =   8421631
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«·€«¡ ”œ«œ"
      PictureAlignment=   3
   End
   Begin VB.Label XONLINE_DOC2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label13"
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
      Left            =   3195
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   630
      Width           =   1455
   End
   Begin VB.Label XSALES_RET 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   630
      Width           =   2985
   End
   Begin VB.Label XONLINE_DOC 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label13"
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
      Left            =   3195
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   135
      Width           =   1455
   End
   Begin VB.Label XPayment_Method 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label13"
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
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   135
      Width           =   2985
   End
End
Attribute VB_Name = "CashOnLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim lExit As Boolean
Dim bSave As Boolean
Public myForm As Form
Private Sub CMD_DELPAY_Click()
If IsDate(xDate_Pay.text) Then
    con.Execute " update file6_20h SET ISNEW = 1, Date_Pay = NULL , bank = NULL WHERE DOC_NO = " & MyParn(SalesFrm.xdoc_no.text)
    Inform "  „ «·€«¡ ”œ«œ «·»Ê‰ "
End If
End Sub
Private Sub CMD_EXIT_Click()
    Unload Me
End Sub
Private Sub CMD_SAVE_Click()
    If Not MYVALID Then Exit Sub
    Dim aInsert(17, 1)
    On Error GoTo myerror
    aInsert(0, 0) = "doc_no"
    aInsert(0, 1) = addstring(SalesFrm.xdoc_no.text)

    aInsert(1, 0) = "name"
    aInsert(1, 1) = addstring(xName.text)

    aInsert(2, 0) = "phone"
    aInsert(2, 1) = addstring(xphone.text)

    aInsert(3, 0) = "address"
    aInsert(3, 1) = addstring(xAddress.text)

    aInsert(4, 0) = "SHIP"
    aInsert(4, 1) = addstring(XSHIP.BoundText)

    aInsert(5, 0) = "CHARGE1"
    aInsert(5, 1) = Val(xcharge1.text)

    aInsert(6, 0) = "CHARGE2"
    aInsert(6, 1) = Val(xcharge2.text)

    aInsert(7, 0) = "ONLINE"
    aInsert(7, 1) = 2
    
    aInsert(8, 0) = "SHIP_NO"
    aInsert(8, 1) = addstring(XSHIP_NO.text)
    
    aInsert(9, 0) = "BANK"
    aInsert(9, 1) = addstring(xbank.BoundText)
    
    aInsert(10, 0) = "cash"
    aInsert(10, 1) = 0
    
    aInsert(11, 0) = "LATE"
    aInsert(11, 1) = Val(SalesFrm.xTotal.text)
    
    aInsert(12, 0) = "PAY"
    aInsert(12, 1) = 0
    
    aInsert(13, 0) = "ISNEW"
    aInsert(13, 1) = 1
    
    aInsert(14, 0) = "rest"
    aInsert(14, 1) = 0
    
    aInsert(15, 0) = "NORET_CHARGE1"
    aInsert(15, 1) = IIf(XNORET.Value, 1, 0)
    
    aInsert(16, 0) = "DATE_REC2"
    aInsert(16, 1) = addDate(XDATE_REC2.text)
    
    aInsert(17, 0) = "ISRETONLINE"
    aInsert(17, 1) = IIf(XISRETONLINE.Value, 1, 0)
    
    con.Execute CreateUpdate(aInsert, "FILE6_20H", " WHERE DOC_NO = " & MyParn(SalesFrm.xdoc_no.text))
    
    aRet = aGetDesca("SELECT DESCA , BRANCH FROM SUBCUST WHERE PHONE = " & MyParn(xphone.text), con)
    If UBound(aRet) = 0 Then
        con.Execute " INSERT INTO [SUBCUST] (desca,phone,f_date,branch,isnew) VALUES ( " & addstring(xName.text) & " , " & addstring(xphone.text) & " , " & addDate(Date) & " , " & addstring(cBranch) & " , 1 ) "
    Else
        con.Execute " update [SUBCUST] set desca = " & addstring(xName.text) & ", isnew = 1 where phone = " & MyParn(xphone.text)
    End If
    con.Execute " update file6_20h set phone = " & addstring(xphone.text) & " where doc_no = " & MyParn(SalesFrm.xdoc_no.text)
    SalesFrm.xphone.Caption = xphone.text
    
    If IsDate(XDATE_REC2.text) Then
        con.Execute " UPDATE FILE6_90H SET DelOrder_Date2 = " & addDate(XDATE_REC2.text) & " WHERE DOC_NO = " & MyParn(XONLINE_DOC2.Caption)
    End If
    Inform " „ Õ›Ÿ »Ì«‰«  «· ”·Ì„ "
    Unload Me
    Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Form_Load()
con.CursorLocation = adUseClient
con.Open strCon

Label12.Visible = lMainShow
xbank.Visible = lMainShow Or (cBranch > "60")

CMD_DELPAY.Visible = lMainShow Or bopt2

xDate_Pay.Locked = Not lMainShow

Set data1.Recordset = myRecordSet("SELECT * FROM SHIP ", con)
Set XSHIP.RowSource = data1
XSHIP.ListField = "Desca"
XSHIP.BoundColumn = "Code"

If cBranch = "00" Then
    Set DATA2.Recordset = myRecordSet("SELECT * FROM FILE5_10 where isship = 1 ORDER BY DESCA ", con)
Else
    Set DATA2.Recordset = myRecordSet("SELECT * FROM FILE5_10 ORDER BY DESCA ", con)
End If
Set xbank.RowSource = DATA2
xbank.ListField = "Desca"
xbank.BoundColumn = "Code"

xTotalQuant.Caption = Myvalue(SalesFrm.xTotalQuant.Caption, "Fixed")
'                           1       2       3       4           5       6       7           8       9      10           11          12              13              14              15          16
aRet = aGetDesca("SELECT PHONE , NAME , CHARGE1 , CHARGE2 , ADDRESS , SHIP , Ship_No , Date_Pay , BANK , note_ret , ONLINE_DOC  , Payment_Method , NORET_CHARGE1 , DATE_REC2  , SALES_RET , ISRETONLINE FROM FILE6_20H WHERE DOC_NO = " & MyParn(SalesFrm.xdoc_no.text), con)
If UBound(aRet) > 0 Then
    xTotal.Caption = Myvalue(SalesFrm.xTotal.text, "Fixed")
    xphone.text = aRet(1) & ""
    xName.text = aRet(2) & ""
    xcharge1.text = Format(Val(aRet(3) & ""), "#0.00")
    xcharge2.text = Format(Val(aRet(4) & ""), "#0.00")
    xTotalCharge.Caption = Format(Val(xTotal.Caption) + Val(xcharge1.text) + Val(xcharge2.text), "#0.00")
    xAddress.text = aRet(5) & ""
    XSHIP.BoundText = aRet(6) & ""
    XSHIP_NO.text = aRet(7) & ""
    xDate_Pay.text = aRet(8) & ""
    xbank.BoundText = aRet(9) & ""
    xnote_ret.text = aRet(10) & ""
    
    XONLINE_DOC.Caption = aRet(11) & ""
    XPayment_Method.Caption = aRet(12) & ""
    XNORET.Value = IIf(aRet(13), 1, 0)
    XDATE_REC2.text = aRet(14) & ""
    XSALES_RET.Caption = aRet(15) & ""
    XISRETONLINE.Value = IIf(aRet(16), 1, 0)
    If XSALES_RET.Caption <> "" Then XONLINE_DOC2.Caption = GetDesca("SELECT ONLINE_DOC FROM FILE6_20H WHERE DOC_NO = " & MyParn(XSALES_RET.Caption), con)
End If
If Not lShowBranch Then
    If IsDate(xDate_Pay.text) And Not lSupperVisor Then cmd_save.Enabled = False
End If
    
    xAddress.Locked = (XONLINE_DOC.Caption <> "") And Not lSupperVisor
    xphone.Locked = (XONLINE_DOC.Caption <> "") And Not lSupperVisor
    xName.Locked = (XONLINE_DOC.Caption <> "") And Not lSupperVisor
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    myForm.lSave = bSave
    Set CashOnLine = Nothing
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) Then
        KeyCode = 0
        SendKeys "{tAB}"
    End If
    If KeyCode = 27 Then Unload Me
End Sub
Private Sub xcharge1_GotFocus()
    xcharge1.SelLength = Len(xcharge1.text)
    xcharge1.BackColor = &H80FFFF
End Sub
Private Sub xcharge1_LostFocus()
    xcharge1.BackColor = &H80000005
    xTotalCharge.Caption = Format(Val(xTotal.Caption) + Val(xcharge1.text) + Val(xcharge2.text), "#0.00")
End Sub
Private Sub xcharge2_GotFocus()
    xcharge2.SelLength = Len(xcharge2.text)
    xcharge2.BackColor = &H80FFFF
End Sub
Private Sub xcharge2_LostFocus()
    xcharge2.BackColor = &H80000005
    xTotalCharge.Caption = Format(Val(xTotal.Caption) + Val(xcharge1.text) + Val(xcharge2.text), "#0.00")
End Sub
Private Sub xDate_Pay_DblClick()
    If Not lMainShow Or cBranch > "60" Then
        If InputBox("", " ”ÃÌ· —ﬁ„ «·»Ê·Ì’…") = XSHIP_NO.text And XSHIP_NO.text <> "" Then
            con.Execute " update file6_20h SET ISNEW = 1, Date_Pay = " & addDate(dSalesDate) & " , bank = " & addstring(xbank.BoundText) & "  WHERE DOC_NO = " & MyParn(SalesFrm.xdoc_no.text)
            Inform "  „  ”ÃÌ· «” ·«„ ﬁÌ„… «·»Ê‰ "
            Unload Me
        End If
    Else
        If xbank.BoundText = "" Then
            MsgBox " √Õ Ì«— »‰ﬂ «·≈Ìœ«⁄ "
            Exit Sub
        End If
        If InputBox("", " ”ÃÌ· —ﬁ„ «·»Ê·Ì’…") = XSHIP_NO.text And XSHIP_NO.text <> "" Then
            xDate_Pay.text = InputBox("", " «—ÌŒ «·≈Ìœ«⁄")
            con.Execute " update file6_20h SET ISNEW = 1 , Date_Pay = " & addDate(xDate_Pay.text) & " WHERE DOC_NO = " & MyParn(SalesFrm.xdoc_no.text)
            Inform "  „  ”ÃÌ· «” ·«„ ﬁÌ„… «·»Ê‰ "
            Unload Me
        End If
    End If
End Sub
Private Sub xdate_ret_Change()
'''''''''''    Dim cNoteRet As String
'''''''''''    If Not IsDate(xDate_Pay.text) Then
'''''''''''        If MsgBox(" „ —›÷ «·≈” ·«„ „‰ «·⁄„Ì· - ”Ê› Ì „ «⁄ »«— »Ê‰ «·„»Ì⁄«  »ﬂ„Ì«  ’›— ", vbYesNo + vbDefaultButton2) = vbYes Then
'''''''''''            If InputBox("", " „ —›÷ «·≈” ·«„ -  ”ÃÌ· —ﬁ„ «·»Ê·Ì’…") = xShip_No.text And xShip_No.text <> "" Then
'''''''''''                cNoteRet = InputBox("", "”»» —›÷ «·«” ·«„")
'''''''''''
'''''''''''                con.Execute " update file6_20h SET discount2 = discount WHERE DOC_NO = " & MyParn(SalesFrm.xdoc_no.text)
'''''''''''                con.Execute " update file6_20h SET Date_ret = " & addDate(dSalesDate) & " , LATE = 0 , discount = 0 , NOTE_RET = " & addstring(cNoteRet) & " WHERE DOC_NO = " & MyParn(SalesFrm.xdoc_no.text)
'''''''''''                con.Execute " update file6_20  SET QUANT2 = QUANT   WHERE DOC_NO = " & MyParn(SalesFrm.xdoc_no.text)
'''''''''''                con.Execute " update file6_20  SET QUANT = 0        WHERE DOC_NO = " & MyParn(SalesFrm.xdoc_no.text)
'''''''''''
'''''''''''                Inform "  „  ”ÃÌ· —›÷ «” ·«„ «·»Ê‰ "
'''''''''''                Unload Me
'''''''''''            End If
'''''''''''        End If
'''''''''''    End If
End Sub
Private Sub xPhone_GotFocus()
    xphone.SelLength = Len(xphone.text)
    xphone.BackColor = &H80FFFF
End Sub
Private Sub XPHONE_LostFocus()
    xphone.BackColor = &H80000005
End Sub
Private Sub XPHONE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If xphone.text <> "" Then
        aRet = aGetDesca("SELECT top 1 [NAME], address FROM file6_20h WHERE PHONE = " & MyParn(xphone.text) & " order by date desc , DOC_NO DESC ", con)
        If UBound(aRet) > 0 Then
            xName.text = aRet(1)
            xAddress.text = aRet(2)
        End If
    End If
End If
End Sub
Private Function MYVALID() As Boolean
If Trim(XSHIP.BoundText) = "" Then
    MsgBox "‘—ﬂ… «·‘Õ‰"
    Exit Function
End If
If Trim(XSHIP_NO.text) = "" Then
    MsgBox "—ﬁ„ »Ê·Ì’… «·‘Õ‰"
    Exit Function
End If
If Trim(xphone.text) = "" Then
    MsgBox "—ﬁ„  ·Ì›Ê‰ «·⁄„Ì· "
    Exit Function
End If
If Trim(xName.text) = "" Then
    MsgBox "≈”„ «·⁄„Ì· "
    Exit Function
End If
If Trim(xAddress.text) = "" Then
    MsgBox "⁄‰Ê«‰ «·⁄„Ì· "
    Exit Function
End If
'If Val(xcharge1.Text) = 0 Then
'    MsgBox "„’«—Ì› «·‘Õ‰"
'    Exit Function
'End If
'If Val(xcharge2.Text) = 0 Then
'    MsgBox "„’«—Ì› »Ê·Ì’… «·”œ«œ"
'    Exit Function
'End If
MYVALID = True
End Function

