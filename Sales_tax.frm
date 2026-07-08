VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form SalesFrm_tax 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15375
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
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   11160
   ScaleWidth      =   15375
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame FRM_CLOSED 
      Height          =   1185
      Left            =   1575
      RightToLeft     =   -1  'True
      TabIndex        =   96
      Top             =   1395
      Width           =   3030
      Begin Threed.SSCommand cmd_closed 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   45
         TabIndex        =   97
         Top             =   630
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   661
         _Version        =   196610
         BackColor       =   8421631
         Caption         =   "≈€·«ﬁ «·„” ‰œ"
      End
      Begin Threed.SSCommand cmd_CLOSEDDATE 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   1530
         TabIndex        =   98
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   661
         _Version        =   196610
         Caption         =   "≈€·«ﬁ › —…"
      End
      Begin Threed.SSCommand cmd_open 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   45
         TabIndex        =   99
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   661
         _Version        =   196610
         BackColor       =   16777152
         Caption         =   "› Õ › —…"
      End
      Begin Threed.SSCheck xclosed 
         Height          =   330
         Left            =   1530
         TabIndex        =   100
         Top             =   675
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   8421631
         PictureMaskColor=   14737632
         Enabled         =   0   'False
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«·„” ‰œ „€·ﬁ"
         Alignment       =   1
         MaskColor       =   8421631
      End
   End
   Begin VB.Frame Frame6 
      Height          =   690
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   89
      Top             =   7290
      Width           =   2445
      Begin VB.CommandButton cmdLasT 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   45
         Picture         =   "Sales_tax.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   93
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   180
         Width           =   510
      End
      Begin VB.CommandButton cmdFirst 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   1845
         Picture         =   "Sales_tax.frx":26DF
         Style           =   1  'Graphical
         TabIndex        =   92
         TabStop         =   0   'False
         ToolTipText     =   "√ŒÌ—"
         Top             =   180
         Width           =   510
      End
      Begin VB.CommandButton cmdNext 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   600
         Picture         =   "Sales_tax.frx":4DB9
         Style           =   1  'Graphical
         TabIndex        =   91
         TabStop         =   0   'False
         ToolTipText     =   "«·”«»ﬁ"
         Top             =   180
         Width           =   555
      End
      Begin VB.CommandButton cmdPrevious 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   1200
         Picture         =   "Sales_tax.frx":738C
         Style           =   1  'Graphical
         TabIndex        =   90
         TabStop         =   0   'False
         ToolTipText     =   "«· «·Ì"
         Top             =   180
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   4140
      Picture         =   "Sales_tax.frx":994C
      RightToLeft     =   -1  'True
      ScaleHeight     =   2400
      ScaleWidth      =   3210
      TabIndex        =   84
      Top             =   7335
      Visible         =   0   'False
      Width           =   3210
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   7380
      Picture         =   "Sales_tax.frx":D023
      RightToLeft     =   -1  'True
      ScaleHeight     =   2355
      ScaleWidth      =   3885
      TabIndex        =   73
      Top             =   7335
      Width           =   3885
      Begin VB.Label xPay 
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
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   1350
         Width           =   1455
      End
      Begin VB.Label xRest 
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
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   1755
         Width           =   1455
      End
      Begin VB.Label xcash 
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
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   81
         Top             =   135
         Width           =   1455
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
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   945
         Width           =   1455
      End
      Begin VB.Label xvisa 
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
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "”œ«œ ‰ﬁœÏ"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   180
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "”œ«œ ›Ì“«"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   585
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "”œ«œ √Ã·"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   990
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œ›Ê⁄"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "«·»«ﬁÏ"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   1845
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   540
      Picture         =   "Sales_tax.frx":10249
      RightToLeft     =   -1  'True
      ScaleHeight     =   780
      ScaleWidth      =   10770
      TabIndex        =   54
      Top             =   9765
      Width           =   10770
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   225
         Width           =   315
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ–› ’‰›"
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
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   225
         Width           =   825
      End
      Begin VB.Shape Shape9 
         BackStyle       =   1  'Opaque
         Height          =   510
         Left            =   90
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   1500
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   480
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   90
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "1-"
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
         Left            =   2160
         TabIndex        =   69
         Top             =   225
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape Shape8 
         BackStyle       =   1  'Opaque
         Height          =   510
         Left            =   1665
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   2835
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   180
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "1+"
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
         Left            =   3150
         TabIndex        =   67
         Top             =   225
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape Shape7 
         BackStyle       =   1  'Opaque
         Height          =   510
         Left            =   2655
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "F5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   225
         Width           =   225
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "”œ«œ ›« Ê—…"
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
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   225
         Width           =   870
      End
      Begin VB.Shape Shape6 
         BackStyle       =   1  'Opaque
         Height          =   510
         Left            =   3645
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   1365
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "F3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   5175
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   225
         Width           =   225
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… Œ’„"
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
         Left            =   5445
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   225
         Width           =   750
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         Height          =   510
         Left            =   5085
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   1230
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "F2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   225
         Width           =   225
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… Œ’„"
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
         Left            =   6795
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   225
         Width           =   780
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         Height          =   510
         Left            =   6390
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   1320
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "F7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   225
         Width           =   225
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "”⁄—"
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
         Left            =   8190
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   225
         Width           =   300
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   510
         Left            =   7785
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   825
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "F6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   225
         Width           =   225
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ"
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
         Left            =   9135
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   225
         Width           =   270
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Esc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   9675
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   225
         Width           =   315
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ï"
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
         Left            =   10035
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   225
         Width           =   510
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         Height          =   510
         Left            =   8685
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   825
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   510
         Left            =   9585
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   1050
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   11385
      Picture         =   "Sales_tax.frx":136E1
      RightToLeft     =   -1  'True
      ScaleHeight     =   3255
      ScaleWidth      =   3885
      TabIndex        =   40
      Top             =   7335
      Width           =   3885
      Begin VB.TextBox xTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   26.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00292952&
         Height          =   540
         Left            =   225
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "123456"
         Top             =   2385
         Width           =   3435
      End
      Begin VB.TextBox xDiscount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   270
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1755
         Width           =   960
      End
      Begin VB.TextBox xRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   1800
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1755
         Width           =   690
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label xNetItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   1350
         Width           =   2220
      End
      Begin VB.Label xDiscItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   945
         Width           =   2220
      End
      Begin VB.Label xTItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   540
         Width           =   2220
      End
      Begin VB.Label xtQuant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   135
         Width           =   2220
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’„ ›« Ê—…"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1845
         Width           =   900
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ï ›« Ê—…"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1395
         Width           =   1035
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’„ √’‰«›"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   990
         Width           =   960
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ï «·ﬁÌ„…"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   585
         Width           =   1005
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ ﬁÿ⁄"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   180
         Width           =   630
      End
   End
   Begin VB.CheckBox chkprint 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "«·€«¡ «·ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2970
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   1035
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Frame Frame5 
      Height          =   780
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   45
      Width           =   9375
      Begin Threed.SSCommand cmdOpen 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   4995
         TabIndex        =   37
         Top             =   135
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1058
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Sales_tax.frx":173DF
         Caption         =   "  »Ê‰«  „› ÊÕ…"
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin Threed.SSCommand SSCommand1 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   7560
         TabIndex        =   38
         Top             =   135
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1058
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Sales_tax.frx":1A854
         Caption         =   "  ÿ»«⁄… «·»Ê‰"
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin Threed.SSCommand SSCommand2 
         CausesValidation=   0   'False
         Height          =   555
         Left            =   90
         TabIndex        =   94
         Top             =   135
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   979
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÿ»«⁄… ≈Ã„«·Ï «·ÌÊ„"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1185
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   1395
      Width           =   1410
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Sales_tax.frx":1D802
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   630
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.CommandButton CmdSave 
         Height          =   465
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Sales_tax.frx":1FD7B
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   9675
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   45
      Width           =   5550
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   4185
         Picture         =   "Sales_tax.frx":201BD
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewInv 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   2835
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Sales_tax.frx":22990
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelInv 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   1485
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Sales_tax.frx":24F3C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Sales_tax.frx":277D6
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1725
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   855
      Width           =   10500
      Begin VB.Frame Frame4 
         Caption         =   "—’Ìœ «·’‰›"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   180
         Width           =   1230
         Begin Threed.SSCommand cmdgo 
            CausesValidation=   0   'False
            Height          =   600
            Left            =   90
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   765
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   1058
            _Version        =   196610
            ForeColor       =   0
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arabic Transparent"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "Sales_tax.frx":27920
            Caption         =   "≈” ⁄·«„"
            Alignment       =   4
            ButtonStyle     =   1
            PictureAlignment=   1
            BevelWidth      =   10
            ShapeSize       =   1
         End
         Begin VB.Label xBalance 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   360
            Width           =   1005
         End
      End
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5310
         MaxLength       =   75
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1260
         Width           =   3840
      End
      Begin VB.CheckBox chkCash 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»Ì⁄ ‰ﬁœÌ"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7560
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   1590
      End
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7560
         MaxLength       =   12
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1590
      End
      Begin VB.TextBox xDate 
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
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   2445
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   540
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo xBox 
         Height          =   315
         Left            =   90
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   900
         Width           =   2445
         _ExtentX        =   4313
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
      Begin MSDataListLib.DataCombo xMan 
         Height          =   330
         Left            =   5310
         TabIndex        =   4
         Top             =   900
         Width           =   3840
         _ExtentX        =   6773
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
      Begin VB.Label xtime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1260
         Width           =   2445
      End
      Begin VB.Label xbalanceitem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1260
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·Êﬁ  :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1260
         Width           =   585
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ«  :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9225
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1260
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·Œ“‰… :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   900
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "«·»«∆⁄ :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9270
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   900
         Width           =   555
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5310
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   540
         Width           =   2220
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   135
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9180
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·›—⁄ :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   540
         Width           =   540
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„Ì· :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9225
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   540
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -2925
      Top             =   3015
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
      Left            =   -1620
      Top             =   1000
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
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   -1890
      Top             =   3915
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   -2430
      Top             =   3555
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
      Left            =   -1620
      Top             =   3465
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
   Begin VB.CheckBox xPrinted 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   1935
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1035
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   285
      Left            =   945
      TabIndex        =   15
      Top             =   10035
      Visible         =   0   'False
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4650
      Left            =   135
      TabIndex        =   8
      Top             =   2610
      Width           =   15090
      _cx             =   26617
      _cy             =   8202
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
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
   Begin VB.Frame Frame9 
      Height          =   600
      Left            =   270
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   6570
      Visible         =   0   'False
      Width           =   3840
      Begin VB.CommandButton cmdTransTo 
         Caption         =   " ÕÊÌ· «·Ì «·»«∆⁄"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   180
         Width           =   1860
      End
      Begin VB.CommandButton cmdTransFrom 
         Caption         =   "”Õ» „‰ «·»«∆⁄"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2025
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   180
         Width           =   1770
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "⁄„Ì· ‰ﬁœÏ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   1575
      RightToLeft     =   -1  'True
      TabIndex        =   85
      Top             =   2610
      Visible         =   0   'False
      Width           =   3120
      Begin VB.TextBox XCODE2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1980
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   86
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label xRate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   88
         Top             =   270
         Width           =   1860
      End
      Begin VB.Label xCust2Desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   630
         Width           =   2940
      End
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3870
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "SalesFrm_tax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_No As String, lSave As Boolean
Dim CardTable As ADODB.Recordset, cFileHeader As String, rdPaid As New ADODB.Recordset
Dim ItemTable As New ADODB.Recordset
Public bRetvalue As Boolean
Dim cDefBox As String, cDefClient As String, cDefClientDesca As String, cDefCasher As String, cDefStore As String, cdefman As String
Dim Search31 As New Search3, search32 As New Search3, bMarket As Boolean
Dim bEdit As Boolean
Dim cFile As String, cFileClient, cCodeDesca As String
Dim con As New ADODB.Connection
Dim formMode, dDateLast As String
Public myPublic As Integer
Const LoadMode = 0, DefineMode = 1
Sub ItemsLookup()
    ItemsLookupAll Me, Search3

End Sub
Private Function myreplace() As Boolean
Dim cSaveMode As String
Dim aInsert(15, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "code"
aInsert(1, 1) = addstring(xCode.Text)

aInsert(2, 0) = "[Date]"
aInsert(2, 1) = addDate(xDate.Text)

aInsert(3, 0) = "store"
aInsert(3, 1) = addstring(xStore.BoundText)

aInsert(4, 0) = "[Notes]"
aInsert(4, 1) = addstring(xNotes.Text)

aInsert(5, 0) = "Rate"
aInsert(5, 1) = Val(xRate.Text)

aInsert(6, 0) = "Discount"
aInsert(6, 1) = Val(xDiscount.Text)

xcash.Caption = Format(Val(xTotal.Text) - Val(xvisa.Caption) - Val(xlate.Caption), "#0.00")
aInsert(7, 0) = "Cash"
aInsert(7, 1) = Val(xcash.Caption)

aInsert(8, 0) = "Box"
aInsert(8, 1) = addstring(xBox.BoundText)

aInsert(9, 0) = "Pay"
aInsert(9, 1) = Val(xPay.Caption)

aInsert(10, 0) = "visa"
aInsert(10, 1) = Val(xvisa.Caption)

aInsert(11, 0) = "late"
aInsert(11, 1) = Val(xlate.Caption)

aInsert(12, 0) = "rest"
aInsert(12, 1) = Val(xRest.Caption)

aInsert(13, 0) = "MAN"
aInsert(13, 1) = addstring(xMan.BoundText)

aInsert(14, 0) = "userName"
aInsert(14, 1) = addstring(sUserName)

aInsert(15, 0) = "code2"
aInsert(15, 1) = addstring(xCode2.Text)


On Error GoTo myerror
con.BeginTrans
If xDoc_No.Tag = DefineMode Then
    xDoc_No.Text = RetZero(Newflag(cFileHeader, "doc_no"), 8)
    aInsert(0, 1) = addstring(xDoc_No.Text)
    con.Execute CreateInsert(aInsert, cFileHeader)
Else
    con.Execute CreateUpdate(aInsert, cFileHeader, " where doc_no = " & addstring(xDoc_No.Text))
End If
myreplaceGrd
con.CommitTrans
myreplace = True
Exit Function
myerror:
prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
xDoc_No.Tag = 1
End Function
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    'Grid1.EditText = Search3.Grid1.TextMatrix(Search3.Grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 1) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    GrdDesc grid1.Row
    If grid1.Row = grid1.Rows - 1 Then
        grid1.TextMatrix(grid1.Rows - 1, 10) = "1"
        grid1.AddItem ""
        MakeSerial
        Grid1_AfterEdit grid1.Row, grid1.Col
        grid1.Select grid1.Rows - 1, 1
    Else
        grid1.TextMatrix(grid1.Row, 10) = "1"
        Grid1_AfterEdit grid1.Row, grid1.Col
        grid1.Select grid1.Row + 1, 1
    End If
    CalcTotals
    
ElseIf ActiveControl.Name = CmdInform.Name Or ActiveControl.Name = cmdOpen.Name Then
    CardTable.Find "DOC_NO = " & MyParn(Search31.grid1.TextMatrix(Search31.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    If ActiveControl.Name = CmdInform Then Search31.Hide Else Unload Search31
    myload
ElseIf ActiveControl.Name = xCode.Name Then
    ActiveControl.Text = search32.grid1.TextMatrix(search32.grid1.Row, 0)
    xCode_LostFocus
    Unload search32
ElseIf ActiveControl.Name = xCode2.Name Then
    ActiveControl.Text = search32.grid1.TextMatrix(search32.grid1.Row, 0)
    xCODE2_LostFocus
    Unload search32

End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub chkprint_Click()
    addSetting "print", chkprint.Value, App.Path & App.Path & "\other.txt"
End Sub

Private Sub cmd_look_Click()

End Sub

Private Sub cmd_closed_Click()
If CardTable!ISCLOSED Then
    con.Execute " update file6_30h set isclosed = 0 where doc_no = " & MyParn(xDoc_No.Text)
Else
    con.Execute " update file6_30h set isclosed = 1 where doc_no = " & MyParn(xDoc_No.Text)
End If
CardTable.Requery
CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub

Private Sub cmd_CLOSEDDATE_Click()
    
    Dim DDate1 As Date, DDate2 As Date
    DDate1 = TurnValue(InputBox("", "≈€·«ﬁ „” ‰œ«  „‰  «—ÌŒ", "1-1"), "", "1-1-2000")
    DDate2 = TurnValue(InputBox("", "≈€·«ﬁ „” ‰œ«  Õ Ï  «—ÌŒ", Format(Date, "DD-MM-YYYY")), "", "1-1-2000")
    If Not IsDate(DDate2) And IsDate(DDate1) Then DDate2 = dDate
    If IsDate(DDate1) And IsDate(DDate2) Then
        If MsgBox("”Ê› Ì „ ≈€·«ﬁ „” ‰œ«  «·› —…", vbOKCancel) = vbOK Then
            con.Execute " update file6_30h set isclosed = 1 where DATE > = " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2), nRec
            MsgBox " „ ≈€·«ﬁ „” ‰œ«  «·› —…" & nRec
            CardTable.Requery
            CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
            If CardTable.EOF Then CardTable.MoveLast
            myload
        End If
    End If

End Sub

Private Sub cmd_open_Click()
    Dim DDate1 As Date, DDate2 As Date
    DDate1 = TurnValue(InputBox("", "› Õ  „” ‰œ«  „‰  «—ÌŒ", ""), "", "1-1-2000")
    DDate2 = TurnValue(InputBox("", "› Õ „” ‰œ«  Õ Ï  «—ÌŒ", ""), "", "1-1-2000")
    If Not IsDate(DDate2) And IsDate(DDate1) Then DDate2 = dDate
    If IsDate(DDate1) And IsDate(DDate2) Then
        If MsgBox("”Ê› Ì „ › Õ „” ‰œ«  «·› —…", vbOKCancel) = vbOK Then
            con.Execute " update file6_30h set isclosed = 0 where DATE > = " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2), nRec
            MsgBox " „ › Õ „” ‰œ«  «·› —…" & nRec
            CardTable.Requery
            CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
            If CardTable.EOF Then CardTable.MoveLast
            myload
        End If
    End If
End Sub


Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    con.BeginTrans
    ' Õ–› «·„” ‰œ
    con.Execute "Delete  From " & cFile & " where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From " & cFileHeader & " where Doc_No = " & MyParn(xDoc_No.Text)
        
    con.CommitTrans
    CardTable.Requery
    
    CmdNewInv_Click
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub CmdFirst_Click()
    CardTable.MoveFirst
    myload
End Sub

Private Sub CmdGo_Click()
    VsModelLook2.Show 1
End Sub
Private Sub CmdInform_Click()
    CardLookup
End Sub
Private Sub CmdLast_Click()
    CardTable.MoveLast
    myload
End Sub
Private Sub CmdNext_Click()
    CardTable.MoveNext
    If CardTable.EOF Then
        CardTable.MovePrevious
    Else
        myload
    End If
End Sub
Private Sub CmdPrevious_Click()
    CardTable.MovePrevious
    If CardTable.BOF Then
        CardTable.MoveNext
    Else
        myload
    End If
End Sub
Private Sub CmdNewInv_Click()
    bAddnew = True
    bEdit = True
    lSave = False
    mydefine
    On Error Resume Next
    xMan.SetFocus
End Sub
Private Sub cmdSave_Click()
    If Not MYVALID Then Exit Sub
    If xCode.Text = "000" Then
        Cashregtax.Show 1
        If Not lSave Then Exit Sub
    Else
        If MsgBox(" ”ÃÌ· „»Ì⁄«  √Ã· ··⁄„Ì·", vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        End If
    End If
    mysave
    On Error Resume Next
    xMan.SetFocus
    Err.Clear
End Sub
Private Sub cmdTransFrom_Click()
    Dim cString As String
    transManfrm.sDate = xDate.Text
    transManfrm.sCaption = "”Õ» „‰ Œ“Ì‰… " & xBox.Text
    transManfrm.sBox1 = xBox.BoundText
    transManfrm.Sbox2 = GetDesca("Select code from file0_50 where type = 1")
    transManfrm.Show 1
End Sub
Private Sub cmdTransTo_Click()
    Dim cString As String
    transManfrm.sDate = xDate.Text
    transManfrm.sCaption = "«Ìœ«⁄ ›Ì Œ“Ì‰… " & xBox.Text
    transManfrm.Sbox2 = xBox.BoundText
    transManfrm.sBox1 = GetDesca("Select code from file0_50 where type = 2")
    transManfrm.Show 1
End Sub
Private Sub CmdUndo_Click()
    CardTable.Requery
    If CardTable.BOF And CardTable.EOF Then
        mydefine
        Exit Sub
    End If
    CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload
End Sub
Private Sub cmdItem_Click()
    Dim bEditLocal As Boolean
    bEditLocal = bEdit: bEdit = True
    ITEMS.Show 1
    bEdit = bEditLocal
End Sub
Private Sub cmdopen_Click()
    CardLookup "PRINTED = 0"
End Sub
Private Sub Command2_Click()
    TDaySal.Show 1
End Sub

Private Sub command31_Click()

End Sub
Private Sub Form_Activate()
    On Error Resume Next
'    Grid1.SetFocus
'    Err.Clear
    xMan.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If xPrinted.Value = 1 Or bEdit = False Then Exit Sub
    If KeyCode = 116 Then
        Grid1_Validate False
        cmdSave_Click
        KeyCode = 0
    ElseIf KeyCode = 115 Then
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
    End If
    If KeyAscii = 27 Then xTotal.SetFocus
    
    Err.Clear
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 117 Then  '  «·ﬂÌ„…
        If grid1.Row <= grid1.Rows - 2 Then
            grid1.Select grid1.Row, 10, grid1.Row, 10
        Else
            grid1.Select 1, 10, 1, 10
        End If
    End If
    
    If KeyCode = 118 Then  '  «·”⁄—
        grid1.Select 1, 11, 1, 11
    End If
    
    If KeyCode = 113 Then  '  ‰”Ì… Œ’„
        xRate.SetFocus
    End If
    If KeyCode = 114 Then  '  ‰”Ì… Œ’„
        xDiscount.SetFocus
    End If
    
    If KeyCode = 115 Then  '  ”œ«œ ›« Ê—…
        cmdSave_Click
    End If
    
    If KeyAscii = 27 Then xTotal.SetFocus   '  «·≈Ã„«·Ï
    
    Err.Clear
End Sub
Private Sub Form_Load()
chkprint.Value = Val(RetSetting("print", App.Path & "\other.txt"))
openCon con
cFileClient = "File3_10"
FRM_CLOSED.Visible = lSupperVisor

Dim aret As Variant
aret = aGetDesca("select code,desca from file3_10 where CODE = '000' ")
cDefClient = "000"
cDefClientDesca = aret(2) & ""

'cDefStore = GetDesca("Select code from file0_40 order by code")
bEdit = cdefman <> ""


cFile = "File6_30"
cFileHeader = "File6_30H"
cMoveName = "„»Ì⁄« "
Me.Caption = "›« Ê—… „»Ì⁄« "
ItemTable.Open "FILE1_10", con, adOpenStatic, adLockReadOnly, adCmdTable
mm = Time
Set CardTable = New ADODB.Recordset
cString = "SELECT " & cFileHeader & ".*,FILE3_10.DESCA AS CLIENTDESCA FROM " & cFileHeader & _
               " INNER JOIN FILE3_10 ON " & cFileHeader & ".Code = FILE3_10.CODE "
If Not bopt3 Then cString = cString & " where FILE6_30H.date = " & DateSq(Date)
If Not bopt3 Then cString = cString & " and  FILE6_30H.box = " & MyParn(cManBox)
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FILE0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

data2.ConnectionString = strCon
data2.RecordSource = "FILE6_25"
Set xMan.RowSource = data2
xMan.ListField = "Desca"
xMan.BoundColumn = "Code"

data4.ConnectionString = strCon
data4.RecordSource = "SELECT * FROM FILE0_50"
Set xBox.RowSource = data4
xBox.ListField = "Desca"
xBox.BoundColumn = "Code"

xMan.BoundText = retDef("file6_25")
With grid1
    .Cols = 13
    .Rows = 2
    .Editable = flexEDKbdMouse
End With

Set grid1.DataSource = DATA3
DATA3.ConnectionString = strCon

If sDoc_No <> "" And Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.Find "doc_no = " & MyParn(sDoc_No), , adSearchForward, adBookmarkFirst
    If Not CardTable.EOF Then
        myload
        Exit Sub
    End If
End If

'xStore.BoundText = retDef("file0_40")
'xMan.BoundText = retDef("file6_25")
'xBox.BoundText = retDef("file0_50")

cdefman = retDef("file6_25")
cDefBox = retDef("file0_50")
cDefStore = retDef("file0_40")

'If Not (CardTable.EOF And CardTable.BOF) Then
    CmdNewInv_Click
'Else
'    mydefine
'    Fixgrd
'    xDoc_No.Text = RetZero("1", 8)
'    xDoc_No.Tag = 1
'End If
'On Error Resume Next
'xMan.SetFocus
If cManBox = "" Then
    MsgBox "«·„” Œœ„ «·Õ«·Ï ·Ì” ·Â… Œ“‰… „”Ã·… ° ·« Ì„ﬂ‰  ”ÃÌ· „»Ì⁄«  ·…"
    Exit Sub
End If
xBox.BoundText = cManBox
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

CardTable.Close
Set CardTable = Nothing

closeCon con

Unload Search3
Unload Search31
Unload search32
Set SalesFrm = Nothing
Err.Clear
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then GrdDesc Row
    myreplaceRow Row, Col
    grid1.Select grid1.Rows - 1, 1
    grid1.ShowCell grid1.Rows - 1, 1
    CalcTotals
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 Then
'    xBalance.Caption = ""
    If Not validRows(OldRow) Then grid1.RemoveItem OldRow
End If
End Sub
Private Sub grid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grid1.EditMaxLength = IIf(Col = 3, 8, 0)
End Sub
Private Sub grid1_EnterCell()
If xPrinted.Value = 1 Or xClosed.Value = -1 Then
    grid1.Editable = flexEDNone
    Exit Sub
End If
'ElseIf (Grid1.Col = 1 And Grid1.TextMatrix(Grid1.Row, Grid1.Cols - 1) <> "") Or (Grid1.Col <> 10 And Grid1.Col <> 12) Then
'    Grid1.Editable = flexEDNone
'ElseIf Grid1.Col = 1 Then
'    Grid1.Editable = flexEDKbdMouse
'ElseIf Grid1.Col = 1 Then
'     Grid1.Editable = flexEDKbdMouse
'Else
'   Grid1.Editable = IIf(Trim(Grid1.TextMatrix(Grid1.Row, 1)) <> "", flexEDKbdMouse, flexEDNone)
'End If
'If Grid1.Col = 10 Then
'    xBalance.Caption = LastBalance(Grid1.TextMatrix(Grid1.Row, 1), xStore.BoundText, con)
'End If

With grid1
    If .Col = 1 Or .Col = 10 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
    .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = vbWhite
    .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = vbYellow
End With
End Sub
Private Sub Grid1_GotFocus()
If grid1.Rows < 2 Then Exit Sub
If grid1.Row = 0 Then
    grid1.Row = 1
    grid1.Col = 1
End If
grid1_EnterCell
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If xPrinted.Value = 1 And Not bopt2 Then Exit Sub

If KeyCode = 13 Then
    KeyCode = 0
    myPos
End If

If KeyCode = 27 Then
    xTotal.SetFocus
    
End If

If KeyCode = 115 Or (KeyCode = 13 And Shift = 2) Then xDiscount.SetFocus
If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 And validRows(grid1.Row) Then
    grid1.AddItem "", grid1.Row
End If
If KeyCode = 112 And xPrinted.Value = 0 And bEdit = True Then
'    Grid1.Row = Grid1.Rows - 1
'    Grid1.Col = 1
'    ItemsLookup
    VsModelLook2.Show 1
End If
End Sub
Private Sub grid1_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If xPrinted.Value = 1 And Not bopt2 Then Exit Sub

If KeyCode = 13 Then
    myPos
End If
End Sub
Private Sub myPos()
'    If grid1.Row = grid1.Rows - 1 Then Exit Sub
    If grid1.Col = 10 And grid1.Row + 1 > grid1.Rows Then
        grid1.Row = grid1.Row + 1
        grid1.Col = IIf(grid1.Row = grid1.Rows - 1, 1, 10)
    ElseIf grid1.Col = 11 Then
        grid1.Row = grid1.Row + 1
        grid1.Col = IIf(grid1.Row = grid1.Rows - 1, 1, 11)
    ElseIf grid1.Col = 1 Then
'        Grid1.Col = 10
        GrdDesc grid1.Row
        If grid1.Rows = grid1.Row + 1 Then grid1.AddItem ""
        grid1.Row = grid1.Row + 1
        grid1.Col = IIf(grid1.Row = grid1.Rows - 1, 1, 10)
     
     End If
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 10 Then xBalance.Caption = LastBalance(grid1.TextMatrix(Row, 1), xStore.BoundText, con)
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
    If Not validRows(grid1.Row) And grid1.Row <> grid1.Rows - 1 Then grid1.RemoveItem grid1.Row
'    xBalance.Caption = ""
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 1 Then
        If grid1.EditText = "" Then
            MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
            Cancel = True
        Else
            If GetDesca("select item from file1_10 where item = " & MyParn(grid1.EditText)) = "" Then
                MsgBox "ﬂÊœ «·’‰› €Ì— ”·Ì„"
                Cancel = True
                Exit Sub
            End If
        End If
    End If
'    If Col = 10 Then
'        If Val(grid1.EditText) < 0 And Not bOpt5 Then
'            MsgBox "€Ì— „”„ÊÕ »«·„— Ã⁄"
'            Cancel = True
'        End If
'    End If
    If Col = 10 Or Col = 11 Then
        If Col = 11 Then
            If Val(grid1.EditText) < Val(grid1.TextMatrix(Row, 14)) Then
                MsgBox "”⁄— €Ì— „”„ÊÕ »…"
                Cancel = True
            End If
        End If
        If Not IsNumeric(grid1.EditText) Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub
Private Sub SSCommand1_Click()
    If Val(xTotal.Text) <> Val(xvisa.Caption) + Val(xcash.Caption) Then
        MsgBox "„—«Ã⁄… ”œ«œ «·»Ê‰ "
        Exit Sub
    End If
    doprint
    SavePrint
End Sub
Private Sub SSCommand2_Click()
On Error GoTo myerror
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
ReDim aHeader(1)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "SELECT  SUPPDESCA, FACTDESCA, SUM(QUANT) AS TQUANT, SUM(TOTAL) AS TTOTAL, BOX FROM         SALES_MODEL "

    cString = cString & " where date = " & DateSq(xDate.Text)
    aHeader(0) = "[" & BetweenString(Format(xDate.Text, "d-m-yyyy"), Format(xDate.Text, "d-m-yyyy")) & "]"

If cManBox <> "" Then
    cString = cString & " AND BOX = " & MyParn(cManBox)
    aHeader(1) = "[" & SalesFrm.xBox.Text & "]"
End If
    
cString = cString & " GROUP BY SUPPDESCA, FACTDESCA, BOX "
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str11 = 1
    temptable!str12 = "≈Ã„«·Ï „Ê—œÌ‰"
    temptable!str3 = " ÌÊ„Ì… " & xDate.Text
    temptable!str1 = !BOX
    temptable!str2 = GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(!BOX))
    temptable!str5 = !SUPPDESCA
    If !SUPPDESCA <> !FACTDESCA Then temptable!str6 = !FACTDESCA
    temptable!val1 = !TQUANT
    temptable!val3 = !TTOTAL
    temptable.Update
    sourcetable.MoveNext
Loop
End With



    cString = "SELECT  MANDESCA, SUM(QUANT) AS TQUANT, SUM(TOTAL) AS TTOTAL, BOX FROM         SALES_MODEL "

    cString = cString & " where date = " & DateSq(xDate.Text)
     aHeader(0) = "[" & BetweenString(Format(xDate.Text, "d-m-yyyy"), Format(xDate.Text, "d-m-yyyy")) & "]"

If cManBox <> "" Then
    cString = cString & " AND BOX = " & MyParn(cManBox)
    aHeader(1) = "[" & SalesFrm.xBox.Text & "]"
End If

sourcetable.Close
cString = cString & " GROUP BY MANDESCA , BOX "
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.Text
    temptable!str1 = !BOX
    
    temptable!str11 = 2
    temptable!str12 = "≈Ã„«·Ï »«∆⁄Ì‰"
    temptable!str2 = GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(!BOX))
    temptable!str5 = !MANDESCA
    
    temptable!val1 = !TQUANT
    temptable!val3 = !TTOTAL
    temptable.Update
    sourcetable.MoveNext
Loop
End With


'
    cString = "SELECT     SUM(T_QUANT) AS T_Q, SUM(t_total) AS T_Item, SUM(discount) AS T_Disc, SUM(CASH) AS T_Cash, SUM(VISA) AS T_Visa FROM         T_SALESDOC"
    cString = cString & " where date = " & DateSq(xDate.Text)
     aHeader(0) = "[" & BetweenString(Format(xDate.Text, "d-m-yyyy"), Format(xDate.Text, "d-m-yyyy")) & "]"
    If cManBox <> "" Then
        cString = cString & " AND BOX = " & MyParn(cManBox)
        aHeader(1) = "[" & SalesFrm.xBox.Text & "]"
    End If

sourcetable.Close
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
With sourcetable
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.Text
    temptable!str1 = SalesFrm.xBox.Text
    
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str2 = SalesFrm.xBox.Text
    temptable!str5 = "⁄œœ „»Ì⁄« "
    
    temptable!val3 = !T_Q
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.Text
    temptable!str1 = SalesFrm.xBox.Text
    
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str2 = SalesFrm.xBox.Text
    temptable!str5 = "ﬁÌ„… «’‰«›"
    
    temptable!val3 = !T_Item
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.Text
    temptable!str1 = SalesFrm.xBox.Text
    
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str2 = SalesFrm.xBox.Text
    temptable!str5 = "ﬁÌ„… «·Œ’„"
    
    temptable!val3 = !T_Disc
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.Text
    temptable!str1 = SalesFrm.xBox.Text
    
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str2 = SalesFrm.xBox.Text
    temptable!str5 = "’«›Ï „»Ì⁄« "
    
    temptable!val3 = Val(!T_Item & "") - Val(!T_Disc & "")
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.Text
    temptable!str1 = SalesFrm.xBox.Text
    
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str2 = SalesFrm.xBox.Text
    temptable!str5 = "”œ«œ ‰ﬁœÏ"
    temptable!val3 = Val(!T_Item & "") - Val(!T_Disc & "") - Val(!T_VISA & "")
    
    temptable.Update
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.Text
    temptable!str1 = SalesFrm.xBox.Text
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str2 = SalesFrm.xBox.Text
    temptable!str5 = "”œ«œ ›Ì“« "
    temptable!val3 = Val(!T_VISA & "")
    temptable.Update
End With
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\TDAY.RPT"
main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.REPORT1.Action = 1
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub xBox_GotFocus()
    ActiveControl.BackColor = &HC0FFFF
End Sub
Private Sub xBox_LostFocus()
    xcash.Enabled = (Trim(xBox.BoundText) <> "")
'  cmdCash.Enabled = (Trim(xBox.BoundText) <> "")
    CalcTotals
    xBox.BackColor = &H80000005
End Sub

Private Sub xCode_DblClick()
    CLIENTLOOKUP
End Sub
Private Sub xCode2_DblClick()
    CLIENT2LOOKUP
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then CLIENTLOOKUP
End Sub
Private Sub xCode_LostFocus()
    xCode.BackColor = &H80000005
    xCodeDesca.Caption = ""
    If xCode.Text = "" Then Exit Sub
    xCode.Text = RetZero(xCode.Text, 3)
    Dim aret
    aret = aGetDesca("select desca,cash from " & cFileClient & " where code = " & MyParn(xCode.Text))
    If UBound(aret) > 0 Then
        xCodeDesca.Caption = aret(1) & ""
        chkCash.Value = IIf(aret(2), 1, 0)
    Else
        xCode.Text = ""
        xCode.SetFocus
    End If
End Sub
Private Sub xCode_Validate(Cancel As Boolean)
    If Trim(xCode.Text) = "" Then Cancel = True
End Sub
Private Sub xCODE2_LostFocus()
    xCode2.BackColor = &H80000005
    xCust2Desca.Caption = ""
    If xCode2.Text = "" Then Exit Sub
    xCode2.Text = xCode2.Text
    Dim aret
    aret = aGetDesca("select desca , DISC from FILE3_20 where code = " & MyParn(xCode2.Text))
    If UBound(aret) > 0 Then
        xCust2Desca.Caption = aret(1) & ""
'        xRate2.Caption = Format(Val(aret(2) & "") / 100, "#.##%")
        xRate2.Caption = aret(2) & ""
    Else
        xCust2Desca.Caption = ""
        xRate2.Caption = ""
    End If
End Sub
Private Sub xDate_Validate(Cancel As Boolean)
    If Not IsDate(xDate.Text) Then Cancel = True
End Sub
Private Sub xDiscount_LostFocus()
    xRate.Text = Format(Val(xDiscount.Text) / Val(xNetItem.Caption) * 100, "#0.00")
    xDiscount.BackColor = &H80000005
    CalcTotals
End Sub
Private Function MYVALID() As Boolean
Dim I As Integer
If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If xBox.BoundText = "" Then
    MsgBox "«·Œ“‰… €Ì— „”Ã·…"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xStore.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
    Exit Function
End If

If xCodeDesca.Caption = "" Then
    MsgBox "·„ Ì „ «œŒ«· ﬂÊœ"
    Exit Function
End If
    
If xMan.BoundText = "" Then
    MsgBox " ”ÃÌ· «·»«∆⁄"
    Exit Function
End If
        

MYVALID = True
End Function
Private Sub myload(Optional bLeaveBal As Boolean = False)
'On Error GoTo myerror
If CardTable!ISCLOSED Then
    cmd_closed.BackColor = &H8080FF
    cmd_closed.Caption = "› Õ „” ‰œ"
    xClosed.Value = ssCBChecked
    xClosed.Visible = True
Else
    cmd_closed.BackColor = vbGreen
    cmd_closed.Caption = "≈€·«ﬁ „” ‰œ"
    xClosed.Value = ssCBUnchecked
    xClosed.Visible = False
End If

xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xStore.BoundText = CardTable!store & ""
xBox.BoundText = CardTable!BOX & ""
xMan.BoundText = CardTable!MAN & ""
xNotes.Text = CardTable!Notes & ""
xCode.Text = CardTable!code & ""
xCodeDesca.Caption = CardTable!ClientDesca & ""
xCode2.Text = CardTable!CODE2 & ""
xCust2Desca.Caption = GetDesca("SELECT DESCA FROM FILE3_20 WHERE CODE = " & MyParn(xCode2.Text)) & ""
xRate2.Caption = GetDesca("SELECT DISC FROM FILE3_20 WHERE CODE = " & MyParn(xCode2.Text)) & ""
chkCash.Value = IIf(CardTable!CASH, 1, 0)
xDiscount.Text = TurnValue(Val(CardTable!DISCOUNT & ""), 0, "")
xRate.Text = TurnValue(Val(CardTable!Rate & ""), 0, "")
xcash.Caption = Format(CardTable!CASH, "#0.00")
xvisa.Caption = TurnValue(CardTable!Visa)
xPay.Caption = Myvalue(CardTable!PAY)
xRest.Caption = Myvalue(CardTable!Rest)
xlate.Caption = Myvalue(CardTable!late)

xPrinted.Value = IIf(CardTable!printed, 1, 0)
xTotal.Enabled = IIf(xPrinted.Value = 0, True, False)
xtime.Caption = Format(CardTable!Time, "hh:nn")
myloadgrd
Handlecontrols LoadMode
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub mydefine()
cmd_closed.BackColor = &H8080FF
cmd_closed.Caption = "› Õ „” ‰œ"
xClosed.Visible = False
xClosed.Value = ssCBUnchecked

xDoc_No.Text = RetZero(Val(Newflag(cFileHeader, "doc_no")), 8)
xDoc_No.Tag = 0
xDate.Text = Format(IIf(Val(Format(Time, "hh")) > 4, Date, DateAdd("d", -1, Date)), "dd-mm-yyyy")
xBalance.Caption = ""
xBox.BoundText = cManBox
xCode.Text = cDefClient
xCodeDesca.Caption = cDefClientDesca
xCode2.Text = ""
xCust2Desca.Caption = ""
xRate2.Caption = ""
'
xStore.BoundText = cDefStore
'xStore.Enabled = False
xMan.BoundText = cdefman

xDiscount.Text = ""
chkCash.Value = 1
xTItem.Caption = ""
xDiscItem.Caption = ""
xNetItem.Caption = ""
xRate.Text = ""
xDiscount.Text = ""
xlate.Caption = ""
xvisa.Caption = ""
xcash.Caption = ""
xPrinted.Value = 0
xMan.BoundText = cdefman
xNotes.Text = ""
xtQuant.Caption = ""
xRest.Caption = ""
xPay.Caption = ""
xtime.Caption = Format(Time, "hh:nn")
grid1.Rows = 1
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
xDoc_No.Tag = 0
xTotal.Text = ""
Fixgrd
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = nMode = LoadMode And bEdit

cmdSave.Enabled = ((bEdit) And xPrinted.Value = 0) Or (bopt2) And (xClosed.Value = 0)
CmdDelInv.Enabled = (nMode = LoadMode And bEdit And xPrinted.Value = 0) Or bopt2 And (xClosed.Value = 0)

cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
xcash.Enabled = (Trim(xBox.BoundText) <> "")
xCode.Enabled = xPrinted.Value = 0 And bEdit = True
End Sub
Private Function retBool(cFieldName) As Boolean
If Not (CardTable.EOF Or CardTable.BOF) Then
    retBool = CardTable(cFieldName)
End If
End Function
Private Sub xDoc_No_LostFocus()
xDoc_No.BackColor = &H80000005
xDoc_No.Text = RetZero(xDoc_No.Text)
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload True
End Sub
Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If xPrinted.Value = 1 And Not bopt2 Then Exit Sub
If KeyCode = 46 And grid1.Row = grid1.Rows - 1 And xPrinted.Value = 0 And bEdit And grid1.Rows > 3 Then
    grid1.Select grid1.Rows - 2, 0
End If
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And xPrinted.Value = 0 And bEdit And grid1.Rows > 3 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        RemoveItem (grid1.Row)
        grid1.Select grid1.Rows - 1, 1
        grid1.ShowCell grid1.Rows - 1, 1
        
        CalcTotals
        'UpdateHeader
        MakeSerial grid1.Row
    End If
End If
'With grid1
'    If .Col = 10 And KeyCode = 22 Then
'        .TextMatrix(.Row, 10) = Val(.TextMatrix(.Row, 10)) + 1
'        KeyCode = 0
'    End If
'End With
'If KeyCode = 27 Then xDate.SetFocus
End Sub
Private Sub GrdDesc(Row)
With grid1
If grid1.Col = 1 And IsNumeric(grid1.TextMatrix(Row, 1)) Then
    ItemTable.Find " item = " & grid1.TextMatrix(Row, 1), , adSearchForward, adBookmarkFirst
    If Not ItemTable.EOF Then
'        xBalance.Caption = Val(GetDesca("select sum([in]-[out] ) from file1_11 where item = " & ItemTable!Item) & "")
        grid1.TextMatrix(Row, 2) = ItemTable!MOSM
        grid1.TextMatrix(Row, 3) = ItemTable!Fact
        grid1.TextMatrix(Row, 4) = ItemTable!supp & ""
        grid1.TextMatrix(Row, 5) = ItemTable!MODELFACT0
        grid1.TextMatrix(Row, 6) = ItemTable!desca
        grid1.TextMatrix(Row, 7) = ItemTable!scal
        grid1.TextMatrix(Row, 8) = ItemTable!Color
        grid1.TextMatrix(Row, 9) = Format(ItemTable!price, "#0.00")
        grid1.TextMatrix(Row, 10) = "1"
        grid1.TextMatrix(Row, 11) = Format(ItemTable!price, "#0.00")
        grid1.TextMatrix(Row, 13) = Format(ItemTable!PRICE2, "#0.00")
        xBalance.Caption = LastBalance(grid1.TextMatrix(Row, 1), xStore.BoundText, con)
        CalcTotals
    End If
End If
End With
End Sub
Private Function CalcTotals(Optional nMode As Integer = 0)
Dim nTotal As Double, nDiscount As Double, nTItem As Double
Dim nDiscItem As Double, nNetItem As Double, nTQuat As Double
Dim nRItem As Double, nTPrice2 As Double
With grid1
For I = 1 To grid1.Rows - 1
    grid1.TextMatrix(I, 13) = Format(Val(grid1.TextMatrix(I, 10)) * Val(grid1.TextMatrix(I, 11)), "#0.00")
    
    nTPrice2 = nTPrice2 + (Val(.TextMatrix(I, 10)) * Val(.TextMatrix(I, 14)))
    nTQuat = nTQuat + Val(.TextMatrix(I, 10))
    nTItem = nTItem + (Val(.TextMatrix(I, 10)) * Val(.TextMatrix(I, 9)))
    nDiscItem = nDiscItem + ((Val(.TextMatrix(I, 9)) - Val(.TextMatrix(I, 11))) * Val(.TextMatrix(I, 10)))
    nNetItem = nNetItem + Val(.TextMatrix(I, 13))
Next
xtQuant.Caption = Format(nTQuat, "#0")
xTItem.Caption = Format(nTItem, "#0.00")
xDiscItem.Caption = Format(nDiscItem, "#0.00")
xNetItem.Caption = Format(nNetItem, "#0.00")
xTotal.Text = Format(Val(xNetItem.Caption) - Val(xDiscount.Text), "#0.00")

'If nMode = 0 Then
'    If xBox.BoundText = "" Then
'        xcash.Caption = Format(0, "Fixed")
'        xlate.Caption = Format(nTotal - Val(xVisa.Caption), "Fixed")
'    Else
'        nCash = IIf(Val(xPay.Caption) >= Val(xTotal.Text), nTotal, Val(xPay.Caption))
'        xcash.Caption = Format(nCash, "Fixed")
'        xlate.Caption = nTotal - (Val(xcash.Caption))
'    End If
'End If
End With
End Function
Private Sub CardLookup(Optional pWhere As String = "")
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT  DOC_NO,DATE , Convert(VARCHAR(10),[DATE],111),  file0_50.desca , Convert(VARCHAR(10),[DATE],108) " & _
                  " FROM  file6_30h inner JOIN file0_50 ON file6_30h.box = file0_50.code where file6_30h.doc_no is not null "
If Not bopt3 Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & " FILE6_30H.DATE = " & DateSq(Date)
End If
If cManBox <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & " FILE6_30H.box = " & MyParn(cManBox)
End If
If pWhere <> "" Then
    Generalarray(1) = Generalarray(1) & " and " & pWhere
End If

Generalarray(2) = "Order by Date , doc_no "
Generalarray(3) = 6000
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ"
listarray(0, 1) = "( Doc_No Like '%cFilter%' or  file0_50.DESCA LIKE '%cFilter%' OR " & _
                  "##date##)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1200

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500


GrdArray(3, 0) = "«·ﬂ«‘Ì— "
GrdArray(3, 1) = 3000

GrdArray(4, 0) = "«·Êﬁ   "
GrdArray(4, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
Search31.Caption = "«” ⁄·«„"
Search31.grid1.FontSize = 10
Search31.Show 1
End Sub

Private Sub xDoc_No_Validate(Cancel As Boolean)
If xDoc_No.Text = "" Then Cancel = True


End Sub
Private Sub xMAN_GotFocus()
    xMan.BackColor = &HC0FFFF
End Sub

Private Sub xMan_LostFocus()
'If Not xDoc_No.Enabled Then UpdateHeader
xMan.BackColor = &H80000005
End Sub
Private Sub xMAN_Validate(Cancel As Boolean)
If IsNumeric(xMan.Text) Then
    xMan.Text = RetZero(xMan.Text, 2)
    xMan.BoundText = xMan.Text
End If
xMan.BoundText = RetZero(xMan.BoundText, 2)
If Not xMan.MatchedWithList Then xMan.BoundText = ""
If Trim(xMan.BoundText) = "" Then
    Cancel = True
    MsgBox " ”ÃÌ· «·»«∆⁄"
End If
End Sub

Private Sub xNotes_LostFocus()
'If Not xDoc_No.Enabled Then UpdateHeader
End Sub
Private Sub xRate_LostFocus()
xRate.BackColor = &H80000005
If Val(xRate.Text) <> 0 Then
    xDiscount.Text = Format(Val(xNetItem.Caption) * (Val(xRate.Text) / 100), "Fixed")
    CalcTotals
End If
'UpdateHeader
End Sub
Private Function RetItemBalance(cItem, cStore, dDate) As Double
If cItem = "" Then Exit Function
movetable.Seek Array(cItem, cStore), adSeekFirstEQ
Do Until movetable.EOF
    If IsNull(movetable!Date) Then Exit Do
    If Trim(movetable!Item) <> cItem Or cStore <> movetable!store Or DateValue(movetable!Date) > DateValue(Format(dDate, "dd-mm-yyyy")) Then Exit Do
    'If Not (movetable!Type = cItemmove And movetable!Doc_Id = xDoc_No.Text) Then
        RetItemBalance = RetItemBalance + TurnValue(movetable!In, Null, 0) - TurnValue(movetable!out, Null, 0)
    'End If
    movetable.MoveNext
Loop
End Function
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For I = 1 To grid1.Rows - 1
    grid1.TextMatrix(I, 0) = I
Next
End Sub
Private Sub Fixgrd()
With grid1
'                0        1           2          3          4           5           6           7       8              9           10          11           12              13              14
.FormatString = "„.|" & "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "”⁄— „” Â·ﬂ|" & "⁄œœ|" & "”⁄— «·»Ì⁄|" & "‰”»… Œ’„|" & "«·≈Ã„«·Ï|" & "”⁄— √Êﬂ«“ÌÊ‰|"
.RowHeight(0) = 1000
.WordWrap = True
.ColHidden(.Cols - 1) = True
.ColHidden(.Cols - 2) = True
.ColWidth(0) = 400
.ColWidth(1) = 1000
.ColWidth(2) = 0
.ColWidth(3) = 1500
.ColWidth(4) = 0
.ColWidth(5) = 1000
.ColWidth(6) = 3800
.ColWidth(7) = 700
.ColWidth(8) = 1000
.ColWidth(9) = 1000
.ColWidth(10) = 1000
.ColWidth(11) = 1000
.ColWidth(12) = 1000
.ColWidth(13) = 1200
.ColFormat(12) = "#.##%"
.ColComboList(3) = StrList("SELECT CODE , DESCA FROM FACT")

.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
.Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter

.ShowCell .Rows - 1, 0

If Not bopt3 And DateValue(Date) <> DateValue(xDate.Text) Then
    grid1.Visible = False
    SSCommand2.Visible = False
    Picture3.Visible = False
    Picture1.Visible = False
Else
    grid1.Visible = True
    SSCommand2.Visible = True
    Picture3.Visible = True
    Picture1.Visible = True
End If
End With
End Sub
Private Sub CLIENTLOOKUP()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From file3_10"
Generalarray(2) = "Order by file3_10.Desca"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load search32
search32.Caption = "«” ⁄·«„"
search32.Show 1
End Sub
Private Sub CLIENT2LOOKUP()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From file3_20"
Generalarray(2) = "Order by file3_20.Desca"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load search32
search32.Caption = "«” ⁄·«„"
search32.Show 1
End Sub

Private Sub xRateDis_Lostfocus()
    xDiscount.Text = Fix((Val(xTotalItem.Caption) * Val(xRateDis.Text) / 100))
End Sub
Private Sub xStore_GotFocus()
    ActiveControl.BackColor = &HC0FFFF
End Sub
Private Sub xstore_LostFocus()
    xStore.BackColor = &H80000005
End Sub
Private Sub xStore_Validate(Cancel As Boolean)
    If Trim(xStore.BoundText) = "" Then Cancel = True
End Sub
Private Sub xTax_LostFocus()
    xTax.BackColor = &H80000005
    CalcTotals
End Sub
Private Function RemoveItem(nRow) As Boolean
On Error GoTo myerror
con.BeginTrans
If grid1.TextMatrix(nRow, grid1.Cols - 1) <> "" Then
    con.Execute "Delete  From " & cFile & " where id = " & grid1.TextMatrix(nRow, grid1.Cols - 1)
    For I = nRow + 1 To grid1.Rows - 2
        con.Execute "update " & cFile & " set [row] = " & (I - 1) & " where id = " & grid1.TextMatrix(nRow, grid1.Cols - 1)
    Next
End If
grid1.RemoveItem nRow
con.CommitTrans
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Function validHeader(Optional bMsg As Boolean = True) As Boolean
If Trim(xDoc_No.Text) = "" Then
    If bMsg Then MsgBox "—ﬁ„ «·›« Ê—… €Ì— „”Ã·"
    Exit Function
End If
If Not IsDate(xDate.Text) Then
    If bMsg Then MsgBox "«· «—ÌŒ €Ì— ’«·Õ «Ê „”Ã·"
    Exit Function
End If

If Trim(xStore.BoundText) = "" Then
    If bMsg Then MsgBox "«·„Œ“‰ €Ì— „”Ã·"
    Exit Function
End If

If Trim(xCode.Text) = "" Then
    If bMsg Then MsgBox "ﬂÊœ «·⁄„Ì· €Ì— „”Ã·"
    Exit Function
End If

validHeader = True
End Function
Private Function validRows(Optional prow = -1, Optional igMsg As Boolean = True, Optional bReqQuant As Boolean = False) As Boolean
For nRow = IIf(prow = -1, 1, prow) To IIf(prow = -1, grid1.Rows - 2, prow)
    If Trim(grid1.TextMatrix(nRow, 1)) = "" Then
        If Not igMsg Then MsgBox "«·’‰› ›Ï «·”ÿ— —ﬁ„ " & nRow & " €Ì— „”Ã· "
        Exit Function
    End If
Next
validRows = True
End Function
Sub myproc2(nDoc_no)
CardTable.Find "Doc_no = " & MyParn(nDoc_no), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
Else
    MsgBox "—ﬁ„ «·›« Ê—… €Ì— ’ÕÌÕ"
    Unload Me
End If
End Sub
Function itemPrice(cItem) As Single
    itemPrice = GetDesca("select PRICE from file1_10 where item = " & MyParn(cItem))
End Function
Private Function myreplaceGrd() As Boolean
Dim aInsert(7, 1)
With grid1
    For I = 1 To grid1.Rows - 2
        If Val(.TextMatrix(I, 11)) = 0 Then
            .TextMatrix(I, 11) = Val(.TextMatrix(I, 9))
        End If
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.Text)
        
        aInsert(1, 0) = "item"
        aInsert(1, 1) = addstring(grid1.TextMatrix(I, 1))
        
        aInsert(2, 0) = "quant"
        aInsert(2, 1) = Val(.TextMatrix(I, 10))

        aInsert(3, 0) = "Price"
        aInsert(3, 1) = Val(.TextMatrix(I, 11))

        aInsert(4, 0) = "Discount"
        aInsert(4, 1) = Val(.TextMatrix(I, 12))

        aInsert(5, 0) = "PRICE_C"
        aInsert(5, 1) = Val(.TextMatrix(I, 9))

        aInsert(6, 0) = "row"
        aInsert(6, 1) = I
        
        aInsert(7, 0) = "COST"
        aInsert(7, 1) = GetDesca("SELECT COST FROM FILE1_10 WHERE ITEM = " & Val(grid1.TextMatrix(I, 1)))

        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, cFile)
        Else
            con.Execute CreateUpdate(aInsert, cFile, " where ID = " & grid1.TextMatrix(I, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Function myreplaceGrdRow(I) As Boolean
Dim aInsert(7, 1)
With grid1
con.BeginTrans
aInsert(0, 0) = "doc_no"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "item"
aInsert(1, 1) = addstring(grid1.TextMatrix(I, 1))

aInsert(2, 0) = "quant"
aInsert(2, 1) = Val(.TextMatrix(I, 10))

aInsert(3, 0) = "Price"
aInsert(3, 1) = Val(.TextMatrix(I, 11))

aInsert(4, 0) = "Discount"
aInsert(4, 1) = Val(.TextMatrix(I, 12))

aInsert(5, 0) = "Price_C"
aInsert(5, 1) = Val(.TextMatrix(I, 9))

aInsert(6, 0) = "row"
aInsert(6, 1) = I

aInsert(7, 0) = "COST"
aInsert(7, 1) = GetDesca("SELECT COST FROM FILE1_10 WHERE ITEM = " & Val(grid1.TextMatrix(I, 1)))

If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
    con.Execute CreateInsert(aInsert, cFile)
Else
    con.Execute CreateUpdate(aInsert, cFile, " where ID = " & grid1.TextMatrix(I, .Cols - 1))
End If
End With
con.CommitTrans
If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then myloadgrd
myreplaceGrdRow = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub myloadgrd()
With grid1
    
    cField1 = "case when file6_30.Discount = 0 then Null else file6_30.Discount end "
'                           0               1               2           3           4               5                       6               7               8           9               10              11              12                  13              14
    cString = "SELECT FILE6_30.ROW, FILE6_30.ITEM, FILE1_10.MOSM, FILE1_10.FACT, FILE1_10.SUPP, FILE1_10.MODELFACT0, FILE1_10.DESCA, FILE1_10.SCAL , FILE1_10.COLOR , FILE6_30.PRICE_C, file6_30.Quant, file6_30.Price, file6_30.DISCOUNT,FILE6_30.TOTAL,FILE1_10.price2,ID " & _
          " FROM FILE6_30 LEFT JOIN FILE1_10 ON FILE6_30.ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(xDoc_No.Text) & " ORDER by FILE6_30.ROW"
    DATA3.RecordSource = cString
    DATA3.Refresh
    grid1.AddItem ""
    MakeSerial
End With
Handlecontrols LoadMode
CalcTotals
Fixgrd
End Sub
Private Sub UpdateHeader()
If Not validHeader Then Exit Sub
myreplace
CardTable.Requery
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub
Private Sub xusername_GotFocus()
xusername.SelStart = 0
xusername.SelLength = Len(xusername.Text)
End Sub
Private Sub xNotes_GotFocus()
xNotes.SelStart = 0
xNotes.SelLength = Len(xNotes.Text)
End Sub
Private Sub xCode_GotFocus()
xCode.SelStart = 0
xCode.SelLength = Len(xCode.Text)
End Sub
Private Sub xDoc_No_GotFocus()
xDoc_No.SelStart = 0
xDoc_No.SelLength = Len(xDoc_No.Text)
End Sub
Private Sub xdate_GotFocus()
xDate.SelStart = 0
xDate.SelLength = Len(xDate.Text)
End Sub
Private Sub xRate_GotFocus()
ActiveControl.BackColor = &HC0FFFF
If xRate.Text = "" Then xRate.Text = Val(xRate2.Caption)
xRate.SelStart = 0
xRate.SelLength = Len(xRate.Text)
End Sub
Private Sub xRateDis_GotFocus()
xRateDis.SelStart = 0
xRateDis.SelLength = Len(xRateDis.Text)
End Sub
Private Sub xDiscount_GotFocus()
ActiveControl.BackColor = &HC0FFFF
xDiscount.SelStart = 0
xDiscount.SelLength = Len(xDiscount.Text)
End Sub
Private Sub xTax_GotFocus()
xTax.SelStart = 0
xTax.SelLength = Len(xTax.Text)
End Sub
Private Function mysave(Optional bEnd As Boolean = True, Optional bPrint As Boolean = True) As Boolean
'lSave = False
If Not MYVALID Then Exit Function
CalcTotals
If Not myreplace Then Exit Function
If bEnd And lSave Then
    SavePrint
    If chkprint.Value = 0 And cManBox <> "" Then
        If doprint Then SavePrint
    Else
        SavePrint
    End If
    
    If xCode.Text = "000" Then
        Inform_OK " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ" & Chr(13) & " «·»«ﬁÏ " & Format(xRest.Caption, "#0.00")
    Else
        Inform_OK " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ" & Chr(13) & "›« Ê—… »Ì⁄ √Ã· ··⁄„Ì· " & xCodeDesca.Caption
    End If
    mydefine
Else
    CardTable.Requery
    CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload
End If
End Function
Private Sub SavePrint()
On Error GoTo myerror
con.BeginTrans
con.Execute "update file6_30h set FILE6_30H.PRINTED = 1 , ISCLOSED = 1  WHERE DOC_NO = " & MyParn(xDoc_No.Text)
con.CommitTrans
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myreplaceRow(Row, Col, Optional pLookup As Boolean = False)
Dim nBalance As Double

If Not validRows(Row) Then Exit Sub
If Row = grid1.Rows - 1 Then
    grid1.AddItem ""
    MakeSerial
End If
If Col = 10 Then
    grid1.TextMatrix(Row, 13) = Format(Val(grid1.TextMatrix(Row, 10)) * Val(grid1.TextMatrix(Row, 11)), "#0.00")
End If

If Col = 11 Then
    nRdisc = Val(grid1.TextMatrix(Row, 9)) - Val(grid1.TextMatrix(Row, 11))
    nRdisc = Format(nRdisc / Val(grid1.TextMatrix(Row, 9)), "#0.00")
    grid1.TextMatrix(Row, 12) = nRdisc
    grid1.TextMatrix(Row, 13) = Format(Val(grid1.TextMatrix(Row, 10)) * Val(grid1.TextMatrix(Row, 11)), "#0.00")
End If
CalcTotals

If xDoc_No.Tag = DefineMode And Row = 1 Then
    mysave False
Else
    myreplaceGrdRow Row
End If
End Sub
Private Sub xTotal_GotFocus()
    xTotal.SelStart = 0
    xTotal.SelLength = Len(xTotal.Text)
End Sub
Private Sub xTotal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(xTotal.Text) < Val(xTPrice.Caption) Then
        MsgBox "«·ﬁÌ„… €Ì— ’ÕÌÕ…"
        xTotal.SetFocus
        Exit Sub
    Else
        xDiscount.Text = Format(Val(xNetItem.Caption) - Val(xTotal.Text), "")
        xRate.Text = Format(xDiscount.Text / Val(xNetItem.Caption) * 100, "#0.00")
        CalcTotals
        Cashregtax.Show 1
        
'       If Not lSave Then Exit Sub
        mysave
        On Error Resume Next
        xMan.SetFocus
        Err.Clear
    
    End If
End If
End Sub
Private Sub xTotal_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        CalcTotals
        xTotal.SelStart = 0
        xTotal.SelLength = Len(xTotal.Text)
    End If
End Sub
Private Function doprint()
Dim datatable As New ADODB.Recordset
datatable.Open "SELECT * FROM FILE6_30 WHERE  DOC_NO = " & MyParn(xDoc_No.Text), con, adOpenKeyset, adLockReadOnly, adCmdText
Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

With datatable
Do While Not datatable.EOF
    temptable.AddNew
    
    temptable!str1 = "‰«œ "
    temptable!str2 = "‰«œ "
    temptable!Date1 = xDate.Text
    temptable!date2 = xtime.Caption
    temptable!str3 = Val(xDoc_No.Text)
    temptable!str7 = xMan.BoundText
    
    temptable!str5 = !Item
    aret = aGetDesca("SELECT desca , fact , modelfact , scal , color FROM FILE1_10 WHERE ITEM = " & !Item)
    If UBound(aret) > 0 Then
        temptable!str6 = aret(1) & " " & DelZero(aret(3)) & " " & GetDesca("select desca from fact where code = " & MyParn(aret(2)))
        If aret(4) <> "---" Then
            temptable!str6 = temptable!str6 & " " & aret(4)
        End If
        If aret(5) <> "---" Then
            temptable!str6 = temptable!str6 & " " & aret(5)
        End If
    
    End If
     
    temptable!val1 = !Quant
    temptable!val2 = !price
    temptable!val3 = !total
    
    temptable!VAL4 = Val(xNetItem.Caption)
    temptable!val5 = Val(xDiscount.Text)
    temptable!Val6 = Val(xTotal.Text)
    temptable!VAL7 = Val(xPay.Caption)
    temptable!Val8 = Val(xRest.Caption)
    temptable!Val10 = Val(xvisa.Caption)
'    temptable!str2 = xusername.Caption
    temptable.Update
    datatable.MoveNext
Loop
End With


contemp.BeginTrans
contemp.CommitTrans

temptable.Requery
REPORT1.ReportFileName = App.Path & "\Reports\sales_tax.rpt"
REPORT1.DataFiles(0) = tempFile
REPORT1.CopiesToPrinter = 1
REPORT1.Action = 1
doprint = True
GoTo closeCon
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
closeCon:
temptable.Close
Set temptable = Nothing
End Function

