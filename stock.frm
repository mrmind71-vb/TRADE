VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form StockFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÊÌ… Ã—œ"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15240
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9855
   ScaleWidth      =   15240
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CheckBox XALLDOC 
      Alignment       =   1  'Right Justify
      Caption         =   "ÿ»«⁄… ﬂ· „” ‰œ«  «·ÌÊ„"
      Height          =   330
      Left            =   3150
      RightToLeft     =   -1  'True
      TabIndex        =   76
      Top             =   810
      Value           =   1  'Checked
      Width           =   2445
   End
   Begin VB.CommandButton POSTOLNY 
      Caption         =   "FIX"
      Height          =   510
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   75
      Top             =   2610
      Width           =   2535
   End
   Begin VB.CommandButton CMD_TOEXECL 
      BackColor       =   &H00FFFFC0&
      Caption         =   " ÕÊÌ· «ﬂ”Ì·"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   15210
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   2745
      Width           =   2040
   End
   Begin VB.Frame Frame12 
      Height          =   600
      Left            =   8910
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Top             =   1710
      Width           =   6270
      Begin MSDataListLib.DataCombo xcode 
         Height          =   315
         Left            =   90
         TabIndex        =   71
         Top             =   180
         Width           =   5280
         _ExtentX        =   9313
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄„Ì· :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   225
         Width           =   570
      End
   End
   Begin VB.CommandButton CMD_FINDBARCODE 
      BackColor       =   &H00C0C0FF&
      Caption         =   "»ÕÀ ⁄‰ »«—ﬂÊœ «Ê „ÊœÌ·"
      Height          =   420
      Left            =   10890
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   2790
      Width           =   2265
   End
   Begin VB.CommandButton FIX_COST 
      Caption         =   "FIX COST"
      Height          =   420
      Left            =   13365
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   2790
      Width           =   1725
   End
   Begin Threed.SSCommand cmd_excel 
      Height          =   465
      Left            =   8865
      TabIndex        =   67
      Top             =   2340
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   13822956
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "stock.frx":0000
      Caption         =   " ÕÊÌ· «ﬂ”Ì· "
      Alignment       =   4
      PictureAlignment=   1
   End
   Begin VB.Frame Frame11 
      Height          =   690
      Left            =   3105
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   1170
      Width           =   3120
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "≈Ã„«·Ï „’«‰⁄"
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   180
         Width           =   825
      End
      Begin Threed.SSCommand CMD_PRINT 
         Height          =   465
         Left            =   990
         TabIndex        =   62
         Top             =   135
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   820
         _Version        =   196610
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
         Picture         =   "stock.frx":27FB
         Caption         =   "ÿ»«⁄… «· ”ÊÌ…"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin VB.Frame Frame9 
      Height          =   2580
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   0
      Width           =   3030
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1125
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   2025
         Width           =   1050
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   63
         Top             =   2025
         Width           =   1005
      End
      Begin VB.CommandButton Command2 
         Caption         =   "√’‰«› ·Â« —’Ìœ"
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
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   135
         Width           =   2940
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   90
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   585
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xfact 
         Height          =   315
         Left            =   90
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   945
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xmosm 
         Height          =   315
         Left            =   90
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1305
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xgroup 
         Height          =   315
         Left            =   90
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1665
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·› —… :"
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   2070
         Width           =   540
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ”„"
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   630
         Width           =   405
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„’‰⁄ "
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   1335
         Width           =   525
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„Ã„Ê⁄…"
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1665
         Width           =   615
      End
   End
   Begin VB.CommandButton cmd_ex 
      BackColor       =   &H0080FF80&
      Caption         =   "«Õ÷«— »Ì«‰«  „‰ «ﬂ”Ì·"
      Height          =   420
      Left            =   10890
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2340
      Width           =   2265
   End
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Caption         =   "Ã—œ »«·»«—ﬂÊœ"
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
      Left            =   13365
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   2340
      Width           =   1725
   End
   Begin VB.CheckBox xClosed 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   330
      Left            =   -45
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   990
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Frame Frame8 
      Height          =   690
      Left            =   3150
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   1890
      Width           =   5640
      Begin VB.TextBox xbalance 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   180
         Width           =   870
      End
      Begin VB.TextBox XTQ1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   180
         Width           =   780
      End
      Begin VB.TextBox XTQ2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3780
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   180
         Width           =   870
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·—’Ìœ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   225
         Width           =   645
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ ﬁÿ⁄"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   225
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ √’‰«›"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   225
         Width           =   855
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   33
      Top             =   9390
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   820
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   12347
            MinWidth        =   12347
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   6255
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   810
      Width           =   1410
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "stock.frx":4C0A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton cmdSave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "stock.frx":7183
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   8820
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   0
      Width           =   6360
      Begin VB.CommandButton CmdDelInv 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   1620
         MaskColor       =   &H00FFFFFF&
         Picture         =   "stock.frx":94E6
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1545
      End
      Begin VB.CommandButton cmdNewInv 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   3195
         MaskColor       =   &H00FFFFFF&
         Picture         =   "stock.frx":BD80
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1545
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   4770
         Picture         =   "stock.frx":E32C
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   135
         Width           =   1545
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "stock.frx":10AFF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1545
      End
   End
   Begin VB.Frame Frame6 
      Height          =   780
      Left            =   3105
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Width           =   5730
      Begin VB.CheckBox XBAL 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ «·—’Ìœ"
         Height          =   330
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   405
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CheckBox XCOST 
         Alignment       =   1  'Right Justify
         Caption         =   "ÿ»«⁄…  ﬁ—Ì— ··Ã—œ "
         Height          =   330
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   135
         Width           =   1635
      End
      Begin Threed.SSCommand STOCK_MODEL 
         Height          =   600
         Left            =   3735
         TabIndex        =   14
         Top             =   135
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1058
         _Version        =   196610
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
         Picture         =   "stock.frx":12F6B
         Caption         =   " ”ÃÌ· „ÊœÌ·«  "
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin Threed.SSCommand CMD_PRINT2 
         Height          =   600
         Left            =   1800
         TabIndex        =   34
         Top             =   135
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1058
         _Version        =   196610
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
         Picture         =   "stock.frx":15773
         Caption         =   "ÿ»«⁄… „” ‰œ «·Ã—œ"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin VB.Frame Frame2 
      Height          =   960
      Left            =   7695
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   765
      Width           =   7485
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "Check2"
         Height          =   195
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   1080
         Width           =   2940
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3150
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   3165
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   540
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4950
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1365
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Width           =   2085
      End
      Begin VB.Label Label4 
         Caption         =   "»Ì«‰ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   585
         Width           =   705
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2220
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„Œ“‰ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   570
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   3240
      Top             =   -45
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   3420
      Top             =   90
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
      Height          =   5295
      Left            =   90
      TabIndex        =   4
      Top             =   3240
      Width           =   15060
      _cx             =   26564
      _cy             =   9340
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   8685
      Width           =   2040
      Begin VB.CommandButton cmdPrevious 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1035
         Picture         =   "stock.frx":17B82
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "”«»ﬁ"
         Top             =   180
         Width           =   465
      End
      Begin VB.CommandButton cmdLast 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         Picture         =   "stock.frx":19DCE
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "√ŒÌ—"
         Top             =   180
         Width           =   465
      End
      Begin VB.CommandButton cmdNext 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   540
         Picture         =   "stock.frx":1C0AB
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "·«Õﬁ"
         Top             =   180
         Width           =   465
      End
      Begin VB.CommandButton cmdFirst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1530
         Picture         =   "stock.frx":1E2D8
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   180
         Width           =   420
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   9945
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   8685
      Width           =   5235
      Begin Threed.SSCommand cmdFix 
         Height          =   510
         Left            =   1620
         TabIndex        =   15
         Top             =   135
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   900
         _Version        =   196610
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
         Picture         =   "stock.frx":2056F
         Caption         =   " «⁄«œ… ÷»ÿ «·Ã—œ"
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin Threed.SSCommand cmdPost 
         Height          =   510
         Left            =   90
         TabIndex        =   16
         Top             =   135
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   900
         _Version        =   196610
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
         Picture         =   "stock.frx":22929
         Caption         =   "  —ÕÌ· «·Ã—œ"
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin Threed.SSCommand cmdUnpost 
         Height          =   510
         Left            =   3510
         TabIndex        =   17
         Top             =   135
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   900
         _Version        =   196610
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
         Picture         =   "stock.frx":24D2B
         Caption         =   " «·€«¡  —ÕÌ·"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin VB.Frame Frame7 
      Height          =   645
      Left            =   3690
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   8685
      Width           =   4110
      Begin VB.CommandButton CMD_FIXALL 
         BackColor       =   &H00FFC0FF&
         Caption         =   "≈⁄«œ…  —ÕÌ· „” ‰œ« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3105
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   135
         Width           =   915
      End
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   510
         Left            =   45
         TabIndex        =   32
         Top             =   135
         Visible         =   0   'False
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   900
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "»ÕÀ ⁄‰ »«—ﬂÊœ ›Ï «·„” ‰œ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   7830
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   8685
      Width           =   2085
      Begin VB.TextBox xfilter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   225
         Width           =   1815
      End
   End
   Begin Threed.SSCommand cmd_closed 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   2205
      TabIndex        =   46
      Top             =   8775
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   979
      _Version        =   196610
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "stock.frx":2714E
      Alignment       =   4
      PictureAlignment=   9
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
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
   Begin MSAdodcLib.Adodc data6 
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
   Begin MSAdodcLib.Adodc data7 
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
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   5640
      Left            =   15210
      TabIndex        =   74
      Top             =   3240
      Width           =   2670
      _cx             =   4710
      _cy             =   9948
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
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
      RightToLeft     =   0   'False
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
   Begin VB.Label xbranch 
      Alignment       =   1  'Right Justify
      Caption         =   "Label7"
      Height          =   285
      Left            =   2205
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   810
      Width           =   240
   End
End
Attribute VB_Name = "StockFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim Search31 As New Search3, search32 As New Search3, oSearchItem As New Search3
Dim CardTable As ADODB.Recordset
Dim cFilter As String
Dim cDefStore As String
Dim formMode, dDateLast As String
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional nRow As Long = -1) As Boolean
Dim aInsert(6, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = DateSq(xDate.text)

aInsert(2, 0) = "store"
aInsert(2, 1) = addstring(xStore.BoundText)

aInsert(3, 0) = "DESCA"
aInsert(3, 1) = addstring(XDESCA.text)

aInsert(4, 0) = "ISDATE"
aInsert(4, 1) = "GETDATE()"

aInsert(5, 0) = "BRANCH"
aInsert(5, 1) = addstring(cBranch)

aInsert(6, 0) = "CODE"
aInsert(6, 1) = addstring(xCode.BoundText)

'con.BeginTrans
'On Error GoTo myerror
If xDoc_No.Tag = DefineMode Then
    xDoc_No.text = RetZero(NewflagBranch6("FILE0_10H", "DOC_NO", cBranch, con), 6)
    aInsert(0, 1) = addstring(xDoc_No.text)
    con.Execute CreateInsert(aInsert, "FILE0_10H")
Else
    con.Execute CreateUpdate(aInsert, "FILE0_10H", " where doc_no = " & addstring(xDoc_No.text))
End If
myreplaceGrd nRow
'con.CommitTrans
myreplace = True
Exit Function
myerror:
prog1.Visible = False
MsgBox Err.Description
'con.RollbackTrans
Err.Clear
End Function
Sub myProc()
'On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0), , 0)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    Dim bNew As Boolean
    bNew = grid1.Row = grid1.Rows - 1
    grid1.TextMatrix(grid1.Row, 0) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    GrdDesc grid1.Row
    grid1.TextMatrix(grid1.Row, 8) = 1
    grid1_AfterEdit grid1.Row, 8
    If Not bNew Then
        Unload oSearchItem
        CellPos 13, grid1.Row, 1
    Else
        CellPos 13, grid1.Row, 8
    End If
Else
    xDoc_No.text = Search31.grid1.TextMatrix(Search31.grid1.Row, 0)
    myUndo
    Unload Search31
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
Unload Search
End Sub

Private Sub Check3_Click()
    addSetting "Check3", Check3.Value, App.Path & "\other.txt"
End Sub

Private Sub cmd_closed_Click()
On Error GoTo myerror
If Not MYVALID Then Exit Sub
Dim StockTable As New ADODB.Recordset
StockTable.Open "SELECT * FROM FILE0_10H WHERE FILE0_10H.date = " & DateSq(xDate.text), con, adOpenStatic, adLockReadOnly, adCmdText
With StockTable
    Do While Not .EOF
        cString = "UPDATE FILE0_10H SET isnew = 1 WHERE DOC_NO = " & MyParn(!DOC_NO)
        con.Execute cString
        .MoveNext
    Loop
End With

myUndo
Inform " O K "
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmd_ex_Click()
    Dim xl As New Excel.Application
    Dim xlsheet As Excel.Worksheet
    Dim xlwbook As Excel.Workbook
    Dim cFileName As String
    Dim nCount As Double, cItem As String, nQTY As Double
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
'           myreplace
            For nRow = 1 To nCount
                If xlsheet.Cells(nRow, 1) <> "" Then
                    cItem = (xlsheet.Cells(nRow, 1))
                    nQTY = Val((xlsheet.Cells(nRow, 2)) & "")
                    cItem = SeekBarcode(grid1.Rows - 1, cItem)
                    If cItem <> "" Then
                        .AddItem ""
                        grid1.TextMatrix(grid1.Rows - 1, 0) = cItem
                        grid1.TextMatrix(grid1.Rows - 1, 8) = nQTY
                        Dim ItemTable As New ADODB.Recordset
                        Set ItemTable = ItemFind(cItem, con)
                        If Not (ItemTable.EOF And ItemTable.BOF) Then
                            grid1.TextMatrix(grid1.Rows - 1, 1) = ItemTable!MOSM
                            grid1.TextMatrix(grid1.Rows - 1, 2) = GetDesca("select desca from fact where code = " & MyParn(ItemTable!Fact), con)
                            grid1.TextMatrix(grid1.Rows - 1, 3) = ItemTable!SUPP & ""
                            grid1.TextMatrix(grid1.Rows - 1, 4) = ItemTable!modelfact0
                            grid1.TextMatrix(grid1.Rows - 1, 5) = ItemTable!DESCA
                            grid1.TextMatrix(grid1.Rows - 1, 6) = ItemTable!color
                            grid1.TextMatrix(grid1.Rows - 1, 7) = ItemTable!scal
                        Else
                            GRID2.AddItem ""
                            GRID2.TextMatrix(GRID2.Rows - 1, 0) = (xlsheet.Cells(nRow, 1))
                            GRID2.TextMatrix(GRID2.Rows - 1, 1) = nQTY
                        End If
                    Else
                        GRID2.AddItem ""
                        GRID2.TextMatrix(GRID2.Rows - 1, 0) = (xlsheet.Cells(nRow, 1))
                        GRID2.TextMatrix(GRID2.Rows - 1, 1) = nQTY
                    End If
                End If
            Next nRow
            End With
            MsgBox "  „ «·≈÷«›… "
        End If
    End If
End Sub

Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub

Private Sub CMD_FINDBARCODE_Click()
    FrmSubStock.Show
End Sub

Private Sub CMD_FIXALL_Click()
Dim cPStore As String, lPost As Boolean
Dim StocklTable As New ADODB.Recordset
If MsgBox(" Ÿ»ÿ ﬂ· „” ‰œ«  «·Ã—œ ", vbOKCancel + vbDefaultButton2) = vbOK Then
    StocklTable.Open "SELECT DATE , STORE  FROM FILE0_10H WHERE CLOSED =1 GROUP BY DATE , STORE ORDER BY STORE , DATE ", con, adOpenStatic, adLockReadOnly, adCmdText
    With StocklTable
    Do While Not .EOF
        Inform " DATE " & !Date & "  - " & !STORE
        xDate.text = !Date
        xStore.BoundText = !STORE
        
        UnPostStock xDate.text, xStore.BoundText
        ReBalStock xDate.text, xStore.BoundText
        PostStock xDate.text, xStore.BoundText
        .MoveNext
    Loop
    End With
End If

'Dim CardTable2 As New ADODB.Recordset
'CardTable2.Open "FR0_10H", con, adOpenStatic, adLockReadOnly, adCmdTable
'Do While Not CardTable2.EOF
'
'    If Not CardTable2!CLOSED Then
'        con.Execute " UPDATE FR0_10H SET CLOSED = 0 WHERE DOC_NO = " & MyParn(CardTable2!doc_no)
'    Else
'        con.Execute " UPDATE FR0_10H SET CLOSED = 0 WHERE DOC_NO = " & MyParn(CardTable2!doc_no)
'        con.Execute " UPDATE FR0_10H SET CLOSED = 1 WHERE DOC_NO = " & MyParn(CardTable2!doc_no)
'    End If
'    Inform " Ã—œ  ÊﬂÌ·«   " & CardTable2!doc_no
'    CardTable2.MoveNext
'Loop
'MsgBox " „ "

End Sub

Private Sub CMD_PRINT_Click()
'On Error GoTo myerror
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
sourcetable.Open "Select FILE0_10.* , file1_10.MODELFACT0 , file1_10.DESCA , file1_10.scal , file1_10.MOSM , file1_10.FACT , file1_10.[SECTION] , FILE1_10.COSTitem  , FILE1_10.COLOR , FACT.DESCA AS FACTNAME , FILE1_10SC.DESCA AS SECEDSCA from FILE0_10 inner join file0_10h on file0_10h.doc_no = file0_10.doc_no INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE0_10.ITEM INNER JOIN FACT ON FACT.CODE = FILE1_10.CODE LEFT JOIN FILE1_10SC ON FILE1_10SC.CODE = FILE1_10.[SECTION] WHERE file0_10h.DATE = " & DateSq(xDate.text) & " AND STORE = " & MyParn(xStore.BoundText), con, adOpenForwardOnly, adLockReadOnly
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
With temptable
Do While Not sourcetable.EOF
    If Round(Val(sourcetable!Differ & ""), 0) <> 0 And Not IsNull(sourcetable!Item) Then
'        ItemTable.Find " ITEM = " & MyParn(sourcetable!Item), , adSearchForward, adBookmarkFirst
'        If Not ItemTable.EOF Then
            .AddNew
            !STR19 = "„” ‰œ Ã—œ " & "  " & XDESCA.text
            !Date1 = xDate.text
            !str11 = sourcetable!Item
            !str1 = TurnValue(sourcetable!modelfact0)
            !str5 = TurnValue(sourcetable!DESCA)
            !str4 = TurnValue(sourcetable!scal)
            !str8 = TurnValue(sourcetable!color)
            !str2 = sourcetable!MOSM
            
            !str3 = TurnValue(sourcetable!FACTNAME)
            !str16 = TurnValue(sourcetable!SECEDSCA)
            
            !str6 = TurnValue(xDoc_No.text, "", Null)
            !STR7 = TurnValue(xStore.text, "", Null)
            !Date1 = xDate.text
            !val3 = sourcetable!RealBal
            !Val8 = sourcetable!Item
            !val2 = sourcetable!ComputerBal
            If sourcetable!Differ > 0 Then
                !VAL6 = Val(sourcetable!Differ)
                !VAL7 = Val(sourcetable!Differ * Val(sourcetable!costITEM & ""))
            Else
                !VAL4 = Val(sourcetable!Differ) * -1
                !val5 = Val(sourcetable!Differ * Val(sourcetable!costITEM & "")) * -1
            End If
            !VAL1 = sourcetable!costITEM
            .Update
'        End If
    End If
    sourcetable.MoveNext
Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
If Check1.Value = 0 Then
    Main.REPORT1.ReportFileName = App.Path & "\Reports\Print_Inv0.rpt"
Else
    Main.REPORT1.ReportFileName = App.Path & "\Reports\Print_Inv0_G.rpt"
End If
Main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear

End Sub
Private Sub Cmd_Print2_Click()
'On Error GoTo myerror
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

''cStr1 = " UPDATE    FILE0_100 SET  cost = Coalesce((SELECT TOP 1 ( FILE7_20.cost_inv ) FROM FILE7_20 INNER JOIN FILE7_20H ON FILE7_20H.DOC_NO = FILE7_20.DOC_NO WHERE FILE0_100.ITEM = FILE7_20.ITEM AND FILE7_20H.DATE <= FILE0_10H.DATE ORDER BY DATE DESC ) ,0) " & _
''        " FROM      FILE0_100 INNER JOIN FILE0_10H ON FILE0_100.DOC_NO = FILE0_10H.DOC_NO "
'con.CommandTimeout = 600
'con.Execute cStr1, nRec


cString = " SELECT  FILE0_100.doc_no, FILE0_100.item, FILE0_100.COST, SUM(FILE0_100.RealBal) AS RealBal , FILE1_10.MODELFACT0, FILE1_10.code, FILE1_10.desca, FILE1_10.FACT, FILE1_10.SCAL, FILE1_10.MOSM , FILE1_10.PRICE , FILE1_10.C_SCAL,FILE4_10.DESCA AS CODEDESCA,FACT.DESCA AS FACTDESCA , FILE1_10.[SECTION] , FILE1_10SC.DESCA AS SECDESCA , FILE1_10.COLOR " & _
        " FROM  FILE0_100 INNER JOIN FILE0_10H ON FILE0_10H.DOC_NO = FILE0_100.DOC_NO INNER JOIN FILE1_10 ON FILE0_100.item = FILE1_10.ITEM  inner JOIN file4_10 ON FILE1_10.CODE = FILE4_10.CODE LEFT JOIN FACT ON FILE1_10.FACT = FACT.CODE LEFT JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE WHERE STORE = " & MyParn(xStore.BoundText)
If XALLDOC.Value = 0 Then cString = cString & " AND FILE0_10H.DOC_NO = " & MyParn(xDoc_No.text)
If XALLDOC.Value <> 0 Then cString = cString & " AND FILE0_10H.DATE = " & DateSq(xDate.text)

cString = cString & " GROUP BY FILE1_10SC.DESCA , FILE1_10.[SECTION], FILE0_100.doc_no, FILE0_100.item, FILE0_100.COST, FILE1_10.MODELFACT0, FILE1_10.code, FILE1_10.desca, FILE1_10.FACT, FILE1_10.SCAL, FILE1_10.MOSM , FILE1_10.PRICE , FILE1_10.C_SCAL,FILE4_10.DESCA ,FACT.DESCA , FILE1_10.COLOR "
sourcetable.Open cString, con, adOpenForwardOnly, adLockReadOnly
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
With temptable
sourcetable.MoveFirst
Do While Not sourcetable.EOF
    If sourcetable!RealBal <> 0 Then
        .AddNew
        !STR7 = "ÿ»«⁄… „” ‰œ «·Ã—œ —ﬁ„ " & xDoc_No.text & "  " & XDESCA.text
        !Date1 = xDate.text
        !str16 = sourcetable!SECDESCA
        !str15 = TurnValue(sourcetable!CodeDesca)
        
        !str5 = sourcetable!CODE
        !str15 = TurnValue(sourcetable!CodeDesca)
        
        !str8 = sourcetable!Item
        !val11 = sourcetable!Item
        !str9 = TurnValue(sourcetable!FACTDESCA)
        !str10 = TurnValue(sourcetable!MOSM)
        !str1 = sourcetable!modelfact0
        !str2 = sourcetable!DESCA
        !str3 = sourcetable!scal
        !str4 = sourcetable!color
'        If xcost.Value = 0 Then
            !VAL4 = sourcetable!RealBal
 '       Else
 '           If XshowBAL.Value <> 0 Then
 '               !VAL4 = RetItemBalance(sourcetable!Item, XSTORE.BoundText, xDate.text, con)
  '          End If
   '     End If
        If bopt2 Then
            !val5 = Val(sourcetable!RealBal & "") * Val(sourcetable!cost & "")
            !VAL6 = Val(sourcetable!cost & "")
        End If
        !Val8 = Val(sourcetable!price & "")
        .Update
        End If
    sourcetable.MoveNext
Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
If XCOST.Value = 0 Then
    Main.REPORT1.ReportFileName = App.Path & "\Reports\Print_Inv2.rpt"
Else
    Main.REPORT1.ReportFileName = App.Path & "\Reports\Print_Inv3.rpt"
End If
Main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear


End Sub
Private Sub cmdDelinv_Click()
    If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        'con.BeginTrans
        con.Execute " Delete  From FILE0_100 where Doc_No = " & MyParn(xDoc_No.text)
        con.Execute " Delete  From FILE0_10 where Doc_No = " & MyParn(xDoc_No.text)
'       con.Execute " Delete  From FILE0_10H where Doc_No = " & MyParn(xDoc_No.Text)
        'con.CommitTrans
        
        AddLod_Data cusername, 2, " Õ–›  " & Me.Caption, con, xDoc_No.text, xDate.text, , XDESCA.text
        
        myUndo
        CmdNewInv_Click
        
        Inform " „ Õ–› «·„” ‰œ »‰Ã«Õ"
    End If
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
    'con.RollbackTrans
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFix_Click()
Dim cDocStock As String

On Error GoTo myerror

AddLod_Data cusername, 1, " ≈«⁄«œ… Ÿ»ÿ Ã—œ  ", con, xDoc_No.text, xDate.text, , XDESCA.text

cDocStock = GetDesca("SELECT MAX(DOC_NO) FROM FILE0_10H WHERE store = " & MyParn(xStore.BoundText) & " and DATE = " & DateSq(xDate.text), con)
con.Execute " DELETE FROM FILE0_10 WHERE FILE0_10.DOC_NO IN (SELECT DOC_NO FROM FILE0_10H WHERE store = " & MyParn(xStore.BoundText) & " and DATE = " & DateSq(xDate.text) & " ) "
con.Execute " INSERT INTO FILE0_10 (doc_no, item  , RealBal , cost)     SELECT  " & addstring(cDocStock) & ", item  , SUM(RealBal) , Coalesce(cost  ,0) FROM  FILE0_100 inner join file0_10h on file0_10h.doc_no = file0_100.doc_no where date = " & DateSq(xDate.text) & " and store = " & MyParn(xStore.BoundText) & " GROUP BY item, cost "
If xCode.BoundText = "" Then
    con.Execute " UPDATE FILE0_10 SET  ComputerBal = Coalesce((select sum([in]-[out]) from file1_11 where file1_11.item = file0_10.item and date < " & DateSq(xDate.text) & " and store = " & MyParn(xStore.BoundText) & " ),0) where doc_no = " & MyParn(cDocStock)
Else
    con.Execute " UPDATE FILE0_10 SET  ComputerBal = Coalesce((select sum([in]-[out]) from file1_11 where file1_11.item = file0_10.item and  file1_11.CUST = " & MyParn(xCode.BoundText) & " and date < " & DateSq(xDate.text) & " and store = " & MyParn(xStore.BoundText) & " ),0) where doc_no = " & MyParn(cDocStock)
End If
con.Execute " UPDATE FILE0_10 SET  Differ = (ComputerBal-RealBal   ) where doc_no = " & MyParn(cDocStock)
If lServerOnLine Then
    cStr1 = " UPDATE    FILE0_10 SET  cost = Coalesce((SELECT TOP 1 ( FILE7_20.cost_inv ) FROM FILE7_20 INNER JOIN FILE7_20H ON FILE7_20H.DOC_NO = FILE7_20.DOC_NO WHERE FILE0_10.ITEM = FILE7_20.ITEM AND FILE7_20H.DATE <= FILE0_10H.DATE ORDER BY DATE DESC ) ,0) " & _
                " FROM  FILE0_10 INNER JOIN FILE0_10H ON FILE0_10.DOC_NO = FILE0_10H.DOC_NO WHERE FILE0_10H.DATE = " & DateSq(xDate.text) & " AND STORE = " & MyParn(xStore.BoundText)
    con.Execute cStr1
Else

    cStr1 = " UPDATE       FILE0_10 SET                cost = Coalesce(FILE1_10.COSTITEM ,0) FROM            FILE1_10 INNER JOIN FILE0_10 ON FILE1_10.ITEM = FILE0_10.item INNER JOIN FILE0_10H ON FILE0_10.doc_no = FILE0_10H.DOC_NO AND FILE0_10.doc_no = FILE0_10H.DOC_NO WHERE FILE0_10H.DATE = " & DateSq(xDate.text) & " AND STORE = " & DateSq(xStore.BoundText)
    con.Execute cStr1
End If
con.Execute " UPDATE    FILE0_10 SET              cost = Coalesce(FILE1_10.COSTITEM ,0) FROM         FILE0_10 INNER JOIN FILE1_10 ON FILE0_10.item = FILE1_10.ITEM where doc_no = " & MyParn(cDocStock)
myUndo
Inform "  „ Ÿ»ÿ «·„” ‰œ "
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(1, 4)
Dim GrdArray(5, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO,DATE, CONVERT(VARCHAR(10),[DATE],111),FILE0_40.DESCA,FILE0_10H.DESCA ,FILE0_10H.CLOSED " & _
                  " FROM (FILE0_10H INNER JOIN FILE0_40 ON FILE0_10H.Store = FILE0_40.CODE) WHERE FILE0_10H.DOC_NO IS NOT NULL "
If lIsBranchStore Then Generalarray(1) = Generalarray(1) & " AND STORE = " & MyParn(cBranchStore)
If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & " AND " & cFilter
End If

Generalarray(2) = "Order by Date DESC , DOC_NO DESC "
Generalarray(3) = 4200
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ-«·»Ì«‰"
listarray(0, 1) = "(@@Doc_No@@6 OR " & _
                  " ##[DATE]## or %%FILE0_10h.DESCA%%)"

listarray(1, 0) = "«·›—⁄"
listarray(1, 1) = "%%FILE0_40.DESCA%%"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "„Œ“‰"
GrdArray(3, 1) = 2000

GrdArray(4, 0) = "«·»Ì«‰"
GrdArray(4, 1) = 2000

GrdArray(5, 0) = " "
GrdArray(5, 1) = 700

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search31
Search31.Caption = "«” ⁄·«„"
Search31.Show 1
End Sub
Private Sub cmdPost_Click()
On Error GoTo myerror
Dim sError As String
'If Not MYVALID Then Exit Sub
AddLod_Data cusername, 1, "  —ÕÌ·  —ÕÌ· Ã—œ  ", con, xDoc_No.text, xDate.text, , XDESCA.text

Dim StockTable As New ADODB.Recordset
StockTable.Open "SELECT * FROM FILE0_10H WHERE store = " & MyParn(xStore.BoundText) & " and FILE0_10H.date = " & DateSq(xDate.text), con, adOpenStatic, adLockReadOnly, adCmdText
With StockTable
    Do While Not .EOF
        Me.Caption = StockTable!DOC_NO
'        cString = " UPDATE FILE0_10H SET FILE0_10H.closed = 0 WHERE DOC_NO = " & MyParn(!DOC_NO)
'        con.Execute cString
'
'        cString = " UPDATE FILE0_10H SET FILE0_10H.closed = 1 , isnew = 1  WHERE DOC_NO = " & MyParn(!DOC_NO)
'        con.Execute cString
        
        
        If Not TransStock(!DOC_NO, 1, con, , sError) Then
            GoTo myerror
        End If
        .MoveNext
    Loop
End With
myUndo
Inform " „  —ÕÌ· «·„” ‰œ »‰Ã«Õ"
Exit Sub
myerror:
MsgBox IIf(sError = "", Err.Description, sError)
Err.Clear
End Sub

Private Sub CmdNewInv_Click()
    myDefine
    On Error Resume Next
    xDoc_No.SetFocus
    Err.Clear
End Sub
Private Sub cmdPrint_Click()
'    doprint
End Sub
Private Sub cmdSave_Click()
'foundOther
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
AddLod_Data cusername, 1, " Õ÷Ÿ " & Me.Caption, con, xDoc_No.text, xDate.text, , XDESCA.text

Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
myUndo
'openCardTable
'myUndo
End Sub
Private Sub CmdUndo_Click()
    myUndo
End Sub
Private Sub cmdunPost_Click()
Dim StockTable As New ADODB.Recordset
Dim sError As String

On Error GoTo myerror

AddLod_Data cusername, 1, " √·€«¡  —ÕÌ· Ã—œ  ", con, xDoc_No.text, xDate.text, , XDESCA.text

StockTable.Open "SELECT * FROM FILE0_10H WHERE FILE0_10H.date = " & DateSq(xDate.text) & " and store = " & MyParn(xStore.BoundText), con, adOpenKeyset, adLockReadOnly, adCmdText

With StockTable
    Do While Not .EOF
        'cString = " UPDATE FILE0_10H SET FILE0_10H.closed = 0 WHERE DOC_NO = " & MyParn(!doc_no)
        'con.Execute cString
        If Not TransStock(!DOC_NO, 0, con, , sError) Then
            GoTo myerror
        End If
        .MoveNext
    Loop
End With
myUndo
Exit Sub
myerror:
MsgBox IIf(sError = "", Err.Description, sError)
Err.Clear
End Sub
Private Sub cmdAddItems_Click()
Dim locTable As New ADODB.Recordset, nRecordCount As Integer

cString = "SELECT FILE1_10.item,file1_10.mosm ,fact.desca as factDesca,file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color,Sum(FILE1_11.[IN] - FILE1_11.[out]) as Balance" & _
          " FROM (FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) INNER JOIN FACT ON FACT.CODE = FILE1_10.FACT"
cString = cString & turn(cString) & " FILE1_11.DATE <= " & DateSq(xDate.text)
cString = cString & turn(cString) & " FILE1_11.STORE = " & MyParn(xStore.BoundText)

cString = cString & " GROUP BY FILE1_10.item,file1_10.mosm ,fact.desca,file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color"
cString = cString & turn(cString, " HAVING ", " AND ") & " Sum(FILE1_11.[IN] - FILE1_11.[out]) <> 0"
locTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText


If Not (locTable.EOF And locTable.BOF) Then
    nRecordCount = locTable.RecordCount
End If
With grid1
    prog1.Visible = True
    prog1.Value = 0
    Do Until locTable.EOF
        If grid1.FindRow(locTable!Item, , 0) = -1 Then
            I = I + 1
            prog1.Value = Round(I / (nRecordCount), 2) * 100
            grid1.TextMatrix(.Rows - 1, 0) = locTable!Item
            grid1.TextMatrix(.Rows - 1, 1) = locTable!MOSM & ""
            grid1.TextMatrix(.Rows - 1, 2) = locTable!FACTDESCA
            grid1.TextMatrix(.Rows - 1, 3) = locTable!SUPP & ""
            grid1.TextMatrix(.Rows - 1, 4) = locTable!modelfact0 & ""
            grid1.TextMatrix(.Rows - 1, 5) = locTable!DESCA & ""
            grid1.TextMatrix(.Rows - 1, 6) = locTable!scal & ""
            grid1.TextMatrix(.Rows - 1, 7) = locTable!color & ""
            grid1.TextMatrix(.Rows - 1, 8) = 0
            grid1.AddItem ""
        End If
        locTable.MoveNext
    Loop
    prog1.Visible = False
End With
End Sub

Private Sub Command1_Click()
Dim locTable As New ADODB.Recordset, nRecordCount As Integer

If Not IsDate(xDate1.text) Then xDate1.text = xDate.text
If Not IsDate(xDate2.text) Then xDate2.text = xDate.text

cString = "SELECT FILE1_10.item,file1_10.mosm ,fact.desca as factDesca,file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color,Sum(FILE1_11.[IN] - FILE1_11.[out]) as Balance" & _
          " FROM (FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) INNER JOIN FACT ON FACT.CODE = FILE1_10.FACT where file1_10.item not in  (select file0_100.item from file0_100 inner join file0_10h on file0_10h.doc_no = file0_100.doc_no where file0_10h.date >= " & DateSq(xDate1.text) & " and date <= " & DateSq(xDate2.text) & " ) "
cString = cString & " and FILE1_11.DATE < " & DateSq(xDate.text)
cString = cString & " and FILE1_11.STORE = " & MyParn(xStore.BoundText)

cString = cString & " GROUP BY FILE1_10.item,file1_10.mosm ,fact.desca,file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color"
cString = cString & turn(cString, " HAVING ", " AND ") & " Sum(FILE1_11.[IN] - FILE1_11.[out]) <> 0"
locTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

If Not (locTable.EOF And locTable.BOF) Then
    nRecordCount = locTable.RecordCount
End If
With grid1
    prog1.Visible = True
    prog1.Value = 0
    Do Until locTable.EOF
        I = I + 1
        prog1.Value = Round(I / (nRecordCount), 2) * 100
        grid1.TextMatrix(.Rows - 1, 0) = locTable!Item
        grid1.TextMatrix(.Rows - 1, 1) = locTable!MOSM & ""
        grid1.TextMatrix(.Rows - 1, 2) = locTable!FACTDESCA
        grid1.TextMatrix(.Rows - 1, 3) = locTable!SUPP & ""
        grid1.TextMatrix(.Rows - 1, 4) = locTable!modelfact0 & ""
        grid1.TextMatrix(.Rows - 1, 5) = locTable!DESCA & ""
        grid1.TextMatrix(.Rows - 1, 6) = locTable!scal & ""
        grid1.TextMatrix(.Rows - 1, 7) = locTable!color & ""
        grid1.TextMatrix(.Rows - 1, 8) = 0
        grid1.AddItem ""
        locTable.MoveNext
    Loop
    prog1.Visible = False
End With
End Sub

Private Sub Command2_Click()
Dim locTable As New ADODB.Recordset, nRecordCount As Integer, cStr1 As String

If Not IsDate(xDate1.text) Then xDate1.text = xDate.text
If Not IsDate(xDate2.text) Then xDate2.text = xDate.text

SubDoStock

cString = "SELECT FILE1_10.item,file1_10.mosm ,fact.desca as factDesca,file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color,Sum(FILE1_11.[IN] - FILE1_11.[out]) as Balance" & _
          " FROM   FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM INNER JOIN FACT ON FACT.CODE = FILE1_10.FACT LEFT OUTER JOIN Q_SUBSTOCK ON FILE1_10.ITEM = Q_SUBSTOCK.item  WHERE (Q_SUBSTOCK.item IS NULL)  "
cString = cString & " and FILE1_11.DATE < " & DateSq(xDate.text)
cString = cString & " and FILE1_11.STORE = " & MyParn(xStore.BoundText)
If xFact.BoundText <> "" Then cString = cString & " and FILE1_10.fact = " & MyParn(xFact.BoundText)
If xMosm.BoundText <> "" Then cString = cString & " and FILE1_10.mosm = " & MyParn(xMosm.BoundText)
If XSECTION.BoundText <> "" Then cString = cString & " and FILE1_10.Section= " & MyParn(XSECTION.BoundText)
If xGroup.BoundText <> "" Then cString = cString & " and FILE1_10.[group]= " & MyParn(xGroup.BoundText)
cString = cString & " GROUP BY FILE1_10.item,file1_10.mosm ,fact.desca,file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color"
cString = cString & turn(cString, " HAVING ", " AND ") & " Sum(FILE1_11.[IN] - FILE1_11.[out]) <> 0"

locTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

If Not (locTable.EOF And locTable.BOF) Then
    nRecordCount = locTable.RecordCount
    locTable.MoveFirst
End If
If nRecordCount = 0 Then
    MsgBox " ·« ÌÊÃœ «’‰«› ··≈ ÷«›… "
    Exit Sub
End If
locTable.MoveFirst
Do Until locTable.EOF
    cString = "Insert into FILE0_100 (doc_no,item,RealBal)" & _
               "Values(" & _
                addstring(xDoc_No.text) & "," & _
                addstring(locTable!Item) & "," & _
                0 & _
               ")"
    con.Execute cString
    locTable.MoveNext
Loop
myloadgrd
'With grid1
''   prog1.Visible = True
'    prog1.Value = 0
'    prog1.Min = 0
'    prog1.Max = nRecordCount
'    Do Until loctable.EOF
'        i = i + 1
'        prog1.Value = i
'        grid1.TextMatrix(.Rows - 1, 0) = loctable!Item
'        grid1.TextMatrix(.Rows - 1, 1) = loctable!MOSM & ""
'        grid1.TextMatrix(.Rows - 1, 2) = loctable!FACTDESCA
'        grid1.TextMatrix(.Rows - 1, 3) = loctable!SUPP & ""
'        grid1.TextMatrix(.Rows - 1, 4) = loctable!MODELFACT0 & ""
'        grid1.TextMatrix(.Rows - 1, 5) = loctable!DESCA & ""
'        grid1.TextMatrix(.Rows - 1, 6) = loctable!scal & ""
'        grid1.TextMatrix(.Rows - 1, 7) = loctable!Color & ""
'        grid1.TextMatrix(.Rows - 1, 8) = 0
'        myreplace .Rows - 1
'        grid1.AddItem ""
'        loctable.MoveNext
'    Loop
'    prog1.Visible = False
'End With
Inform "  „ «÷«›… «·„ÊœÌ·«  "
End Sub

Private Sub FIX_COST_Click()
    cStr1 = " UPDATE    FILE0_10 SET  cost = Coalesce((SELECT TOP 1 ( FILE7_20.cost_inv ) FROM FILE7_20 INNER JOIN FILE7_20H ON FILE7_20H.DOC_NO = FILE7_20.DOC_NO WHERE FILE0_10.ITEM = FILE7_20.ITEM AND FILE7_20H.DATE <= FILE0_10H.DATE ORDER BY DATE DESC ) ,0) " & _
            " FROM      FILE0_10 INNER JOIN FILE0_10H ON FILE0_10.DOC_NO = FILE0_10H.DOC_NO"
    con.CommandTimeout = 600
    con.Execute cStr1, nRec
    Inform "  „ Ÿ»ÿ  ”ÊÌ… «·Ã—œ " & nRec

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
If KeyAscii = 43 Then
    KeyAscii = 0
    If grid1.Row > 1 Then
        grid1.Row = grid1.Row - 1
        grid1.Col = 8
    End If
End If
End Sub
Private Sub Form_Load()
bedit = True
openCon con

FIX_COST.Visible = (cBranch = "00")
Check3.Value = Val(RetSetting("Check3", App.Path & "\other.txt"))

Set DATA3.Recordset = myRecordSet("SELECT * FROM FILE1_10sc ORDER BY code ", con)
Set XSECTION.RowSource = DATA3
XSECTION.ListField = "Desca"
XSECTION.BoundColumn = "Code"

Set DATA4.Recordset = myRecordSet("SELECT * FROM fact ORDER BY code ", con)
Set xFact.RowSource = DATA4
xFact.ListField = "Desca"
xFact.BoundColumn = "Code"

Set DATA5.Recordset = myRecordSet("SELECT * FROM mosm ORDER BY date desc ", con)
Set xMosm.RowSource = DATA5
xMosm.ListField = "Desca"
xMosm.BoundColumn = "mosm"

Set DATA7.Recordset = myRecordSet("SELECT * from file3_10 WHERE ISONEST = 1  ORDER BY code ", con)
Set xCode.RowSource = DATA7
xCode.ListField = "Desca"
xCode.BoundColumn = "Code"

Set DATA6.Recordset = myRecordSet("SELECT * from file1_50 ORDER BY desca ", con)
Set xGroup.RowSource = DATA6
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

Set DATA2.Recordset = myRecordSet("SELECT * from file0_40 ORDER BY code ", con)
Set xStore.RowSource = DATA2
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

If cBranch <> "00" Then
    xStore.BoundText = cBranchStore
    xStore.Enabled = False
End If


Set grid1.DataSource = data1

myUndo


End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Unload Search3
Unload Search31
If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
Err.Clear
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim bNew As Boolean
If grid1.Col = 0 Then
    GrdDesc grid1.Row
End If

If Not validRow(Row) Then Exit Sub
If Row = grid1.Rows - 1 Then
    myaddItem
    bNew = True
End If
CalcTotals

With grid1
If myreplace(Row) Then
    If xDoc_No.Tag = DefineMode Then xDoc_No.Tag = LoadMode
Else
    bNew = False
End If
myloadgrd
If bNew Then
    If Check3.Value = 0 Then
        grid1.Row = grid1.Row
        grid1.Col = 8
        grid1.ShowCell grid1.Row, 8
    Else
        grid1.Row = grid1.Rows - 1
        grid1.ShowCell grid1.Rows - 1, 0
        grid1.Col = 0
    End If
End If

End With
End Sub

Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then grid1.RemoveItem OldRow
End If
End Sub

Private Sub grid1_EnterCell()
With grid1
If (grid1.Col = 0 Or grid1.Col = 8) And cmdSave.Enabled Then
''    .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = vbWhite
    .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = vbYellow
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End With
End Sub
Private Sub Grid1_GotFocus()
With grid1
    If grid1.Row <= 1 Then
    .Select 1, 0, 1, 0
    .ShowCell 1, 0
    End If
End With
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 Then grid1.AddItem "", grid1.Row
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'If Col = 2 And Trim(grid1.TextMatrix(Row, 3)) = "" Then
    '    nBalance = RetItemBalance(grid1.TextMatrix(Row, 0), xstore1.BoundText, DateAdd("D", -1, xDate.Text), con)
    '    grid1.TextMatrix(Row, 9) = nBalance
    'End If
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
    If (Not validRow(grid1.Row)) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then grid1.RemoveItem grid1.Row
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    With grid1
    If .EditText <> "" Then
        If UCase(Mid(.EditText, 1, 1)) = "J" Then
            cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE2 = " & MyParn(.EditText), con)
        ElseIf UCase(Mid(.EditText, 1, 1)) = "S" Then
            cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE2 = " & MyParn(.EditText), con)
        ElseIf Len(.EditText) >= 10 Then
            cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE13 = " & MyParn(.EditText) & " OR BARCODE = " & MyParn(.EditText) & " OR BARCODE_GS1 = " & MyParn(.EditText), con)
            If cMyItem = "" Then cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE = " & MyParn(.EditText), con)
            
        End If
        If cMyItem <> "" Then
            .EditText = Val(cMyItem)
        End If
    End If
    If Len(.EditText) > 10 Then .EditText = ""
    End With
'    If Not IsNumeric(grid1.EditText) Then
'        Inform "»«—ﬂÊœ «·’‰› €Ì— ”·Ì„"
'        Cancel = True
'        Exit Sub
'    End If
    If Not validItem(grid1.EditText, con) Then
        Inform "»«—ﬂÊœ «·’‰› €Ì— „ÊÃÊœ"
        Cancel = True
    End If
'    nFound = FoundOtheritem(Row, 0, Trim(grid1.EditText))
'    If nFound <> -1 Then
'        MsgBox "«·’‰› „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & nFound
'        Cancel = True
'    End If
End If
'If Col = 8 Then
'    If Val(grid1.EditText) >= Val(grid1.TextMatrix(Row, 9)) Then
'        MsgBox "«·—’Ìœ ·« Ì”„Õ"
'    End If
'End If
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub POSTOLNY_Click()
Dim cPStore As String, lPost As Boolean
Dim StocklTable As New ADODB.Recordset
If MsgBox(" Ÿ»ÿ ﬂ· „” ‰œ«  «·Ã—œ ", vbOKCancel + vbDefaultButton2) = vbOK Then
    StocklTable.Open "SELECT DOC_NO , DATE , STORE  FROM FILE0_10H WHERE CLOSED = 1 ORDER BY DOC_NO ", con, adOpenStatic, adLockReadOnly, adCmdText
    With StocklTable
    .MoveFirst
    Do While Not .EOF
        Me.Caption = " DATE " & !Date & "  - " & !STORE & "  - " & !DOC_NO
        xDate.text = !Date
        xStore.BoundText = !STORE
        con.Execute " UPDATE FILE0_10H SET CLOSED =  0 WHERE DOC_NO = " & MyParn(!DOC_NO)
con.Execute " UPDATE FILE0_10H SET CLOSED =1 WHERE DOC_NO = " & MyParn(StocklTable!DOC_NO)
'        UnPostStock xDate.text, xStore.BoundText
'        ReBalStock xDate.text, xStore.BoundText
'        PostStock xDate.text, xStore.BoundText
        .MoveNext
    Loop
    End With
End If

If MsgBox(" Ÿ»ÿ ﬂ· „” ‰œ«  «·Ã—œ  ÊﬂÌ·«  ", vbOKCancel + vbDefaultButton2) = vbOK Then
    StocklTable.Open "SELECT DOC_NO , DATE , STORE  FROM FR0_10H WHERE CLOSED = 1 ORDER BY DOC_NO ", con, adOpenStatic, adLockReadOnly, adCmdText
    With StocklTable
    .MoveFirst
    Do While Not .EOF
        Me.Caption = " DATE " & !Date & "  - " & !STORE & "  - " & !DOC_NO
        xDate.text = !Date
        xStore.BoundText = !STORE
        con.Execute " UPDATE FILE0_10H SET CLOSED =  0 WHERE DOC_NO = " & MyParn(!DOC_NO)
        con.Execute " UPDATE FILE0_10H SET CLOSED =1 WHERE DOC_NO = " & MyParn(StocklTable!DOC_NO)
'        UnPostStock xDate.text, xStore.BoundText
'        ReBalStock xDate.text, xStore.BoundText
'        PostStock xDate.text, xStore.BoundText
        .MoveNext
    Loop
    End With
End If

End Sub

Private Sub STOCK_MODEL_Click()
If Not MYVALID2 Then Exit Sub
StockModelfrm.Show 1

myUndo
End Sub
Private Sub xdate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
End Sub

Private Sub xDate_Validate(Cancel As Boolean)
myValidDate xDate
End Sub

Private Sub xDoc_No_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CmdInform_Click
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If xDoc_No.text = "" Then
    If Not bIsgMsg Then Inform "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    If Not bIsMsg Then Inform "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
If xStore.BoundText = "" Then
    If Not bIsMsg Then Inform "·„ Ì „ «œŒ«· «·„Œ“‰ "
    Exit Function
End If

If grid1.Rows < 3 Then
    If Not bIsMsg Then Inform "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
    Exit Function
End If

With grid1
For I = 1 To .Rows - 2
    If Not IsNumeric(.TextMatrix(I, 0)) Then
        If Not validItem(.TextMatrix(I, 0), con) Then
            MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
            Exit Function
        End If
    End If
Next
End With
MYVALID = True
End Function
Private Function MYVALID2() As Boolean
If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
If xStore.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
    Exit Function
End If

MYVALID2 = True
End Function
Private Sub myload()
xDoc_No.text = CardTable!DOC_NO
xDate.text = Format(CardTable!Date, "dd-mm-yyyy")
xStore.BoundText = CardTable!STORE
XDESCA.text = CardTable!DESCA & ""
xCode.BoundText = CardTable!CODE & ""

xClosed.Value = IIf(CardTable!closed, 1, 0)
xDoc_No.Enabled = False
myloadgrd
Handlecontrols LoadMode
CalcTotals
'cmdSave.Enabled = False

End Sub
Private Sub myDefine()
xDoc_No.Enabled = True

xDoc_No.text = RetZero(NewflagBranch6("FILE0_10H", "DOC_NO", cBranch, con), 6)
xDate.text = Format(Date, "dd-mm-yyyy")
xStore.BoundText = cDefStore
xCode.BoundText = ""

XDESCA.text = ""
xClosed.Value = 0
xStore.Enabled = False
xStore.BoundText = cBranchStore
'StatusBar1.Panels(1).Text = ""
'xTotal.Caption = ""
grid1.Rows = 1
grid1.AddItem ""
Handlecontrols DefineMode
STOCK_MODEL.Enabled = True
Fixgrd
End Sub
Private Sub Handlecontrols(nMode)
Dim bClosed As Boolean
cmdNewInv.Enabled = nMode = LoadMode
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
bClosed = xClosed.Value = 1
cmdPost.Enabled = (Not bClosed) And nMode = LoadMode

cmdUnpost.Enabled = bClosed And bedit And nMode = LoadMode
cmdFix.Enabled = (Not bClosed) And bedit And nMode = LoadMode
cmdSave.Enabled = (nMode = DefineMode Or Not bClosed)
CmdDelInv.Enabled = (Not bClosed) And nMode = LoadMode
Command2.Enabled = (Not bClosed) And nMode = LoadMode


xStore.Enabled = cDefStore = ""
STOCK_MODEL.Enabled = (Not bClosed)
xDoc_No.Tag = nMode
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No
If xDoc_No.text = "" Then Exit Sub
xDoc_No.text = RetZero(xDoc_No.text)
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 0 Then
    ItemsLookupAll Me, oSearchItem
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            'con.BeginTrans
            con.Execute "delete from FILE0_100 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            'con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
    End If
ElseIf KeyCode = 13 Then
'   CellPos KeyCode, grid1.Row, grid1.Col
'ElseIf KeyCode = 45 And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And validRow(grid1.Row) Then
'    grid1.AddItem "", grid1.Row
End If
Exit Sub
myerror:
MsgBox Err.Description
'con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 112 And Col = 0 Then
     ItemsLookupAll Me, oSearchItem
ElseIf KeyCode = 13 Then
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub GrdDesc(Row)
Dim nBalance As Double
'On Error GoTo myerror
If Not IsNumeric(grid1.TextMatrix(Row, 0)) Then Exit Sub
nBalance = RetItemBalance(grid1.TextMatrix(Row, 0), xStore.BoundText, xDate.text, con, xCode.BoundText)

xBalance.text = Format(nBalance, "#0")
'grid1.TextMatrix(Row, 9) = nBalance
Dim ItemTable As New ADODB.Recordset
Set ItemTable = ItemFind(grid1.TextMatrix(Row, 0), con)
With grid1
    If Not (ItemTable.EOF And ItemTable.BOF) Then
        If ItemTable!ISSTOP Then
            MsgBox "’‰› „ Êﬁ›"
            Exit Sub
        End If
        
        grid1.TextMatrix(grid1.Row, 1) = ItemTable!MOSM
        grid1.TextMatrix(grid1.Row, 2) = GetDesca("select desca from fact where code = " & MyParn(ItemTable!Fact), con)
        grid1.TextMatrix(grid1.Row, 3) = ItemTable!SUPP & ""
        grid1.TextMatrix(grid1.Row, 4) = ItemTable!modelfact0
        grid1.TextMatrix(grid1.Row, 5) = ItemTable!DESCA
        grid1.TextMatrix(grid1.Row, 6) = ItemTable!color
        grid1.TextMatrix(grid1.Row, 7) = ItemTable!scal
        grid1.TextMatrix(grid1.Row, 8) = IIf(grid1.TextMatrix(grid1.Row, 8) = "", 1, grid1.TextMatrix(grid1.Row, 8))
        
    End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function CalcTotals()
Dim nCount As Double
Dim nTotalQuant As Double, nTotalCost As Double
With grid1
'For i = 1 To grid1.Rows - 2
''   nCount = nCount + 1
''    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(I, 8))
'Next
'Me.StatusBar1.Panels(1).Text = IIf(nTotalQuant <> 0, "≈Ã„«·Ì «·ﬂ„Ì… : " & nTotalQuant, "")
'xCount.Caption = "⁄œœ «·«’‰«› " & Format(nCount, "#0")
End With
End Function
Private Function FoundOtherRow(nRow, nCol) As Integer
FoundOtherRow = -1
For I = 1 To grid1.Rows - 2
    If I <> nRow Then
        If Trim(grid1.TextMatrix(I, nCol)) = Trim(grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = I
            Exit Function
        End If
    End If
Next
End Function

Private Function FoundOtheritem(nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For I = 1 To grid1.Rows - 2
    If I <> nRow Then
        If Trim(grid1.TextMatrix(I, nCol)) = nValue Then
            FoundOtheritem = I
            Exit Function
        End If
    End If
Next
End Function
Private Sub Fixgrd()
Dim n1 As Double, n2 As Double
With grid1
'                   0            1      2               3        4           5       6           7          8            9
.FormatString = "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "⁄œœ Ã—œ|" & "„|"
.WordWrap = True

.RowHeight(0) = 1000

.ColWidth(0) = 1000
.ColWidth(1) = 1000
.ColWidth(2) = 2000
.ColWidth(3) = 1000
.ColWidth(4) = 1500
.ColWidth(5) = 3200
.ColWidth(6) = 1000
.ColWidth(7) = 1000
.ColWidth(8) = 900
.ColWidth(9) = 1000
.ColWidth(10) = 1000
'.ColHidden(9) = True
'.ColHidden(10) = True
.ColHidden(.Cols - 1) = True

For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
n1 = .Rows - 2
For nRow = 1 To .Rows - 1
    .TextMatrix(nRow, 9) = nRow
    n2 = n2 + Val(.TextMatrix(nRow, 8))
Next nRow
XTQ1.text = n1
XTQ2.text = n2
End With
End Sub
Private Sub myreplaceGrd(nRow)
Dim aInsert(2, 1)
With grid1
    For I = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
        If .TextMatrix(I, 0) <> "" Then
        
            aInsert(0, 0) = "doc_no"
            aInsert(0, 1) = addstring(xDoc_No.text)
            
            aInsert(1, 0) = "item"
            aInsert(1, 1) = addstring(grid1.TextMatrix(I, 0))
            
            aInsert(2, 0) = "RealBal"
            aInsert(2, 1) = Val(.TextMatrix(I, 8))
                                                                                            
            If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
                con.Execute CreateInsert(aInsert, "FILE0_100")
            Else
                con.Execute CreateUpdate(aInsert, "FILE0_100", " where ID = " & grid1.TextMatrix(I, .Cols - 1))
            End If
        End If
    Next
End With
End Sub
Private Sub myloadgrd()
Dim cString As String
cString = "SELECT FILE0_100.item , file1_10.mosm , fact.desca , file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color ,  FILE0_100.RealBal, ' ' as n1 , FILE0_100.ID from (FILE0_100 inner join file1_10 on file1_10.item = FILE0_100.item) inner join fact on fact.code = file1_10.fact"
cString = cString & turn(cString) & " DOC_NO = " & MyParn(xDoc_No.text)
cString = cString & " ORDER BY ID"
Set data1.Recordset = myRecordSet(cString, con)
grid1.AddItem ""
Fixgrd
End Sub
Private Function MYVALIDfix() As Boolean
If foundOther Then Exit Function
If xDoc_No.text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xStore.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰"
    Exit Function
End If

If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
    Exit Function
End If


With grid1
For I = 1 To .Rows - 2
    If .TextMatrix(I, 0) = "" Then
        .Select I, 0, I, grid1.Cols - 1
        MsgBox "ﬂÊœ «·’‰› €Ì— „ÊÃÊœ"
        Exit Function
    Else
        If GetDesca("SELECT ITEM FROM FILE1_10 WHERE ITEM = " & MyParn(.TextMatrix(I, 0)), con) = "" Then
            .Select I, 0, I, 2
            MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
            Exit Function
        End If
    End If
Next
End With
MYVALIDfix = True
End Function
Private Sub myUndo()
If IsNumeric(xDoc_No.text) Then
    openCardTable xDoc_No.text
    If Not CardTable.EOF Then
        myload
        Exit Sub
    End If
End If
openCardTable , "<"
If CardTable.EOF Then myDefine Else myload
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub xfilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FilterGrd grid1, xfilter.text, 0
    End If
End Sub
Private Sub SubReplaceItem(pItem, nReal, nComputer, nDiffer)
Dim aInsert(4, 1)
With grid1
    aInsert(0, 0) = "doc_no"
    aInsert(0, 1) = addstring(xDoc_No.text)
        
    aInsert(1, 0) = "item"
    aInsert(1, 1) = addstring(pItem)
        
    aInsert(2, 0) = "RealBal"
    aInsert(2, 1) = nReal
        
    aInsert(3, 0) = "ComputerBal"
    aInsert(3, 1) = nComputer
        
    aInsert(4, 0) = "Differ"
    aInsert(4, 1) = nDiffer
    con.Execute CreateInsert(aInsert, "FILE0_10")
End With
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = True) As Boolean
With grid1
If Not IsNumeric(.TextMatrix(Row, 0)) Then Exit Function
'If Not validItem(.TextMatrix(Row, 0), con) Then Exit Function
validRow = True
End With
End Function
Private Sub myaddItem()
grid1.Row = grid1.Row
grid1.Col = 8
grid1.Select grid1.Row, 8

grid1.AddItem ""

End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col = 0 Then
    If Not IsNumeric(grid1.TextMatrix(Row, 0)) Then Exit Sub
    If Row = grid1.Rows - 1 Then
        grid1.Col = 8
        grid1.ShowCell Row, 8
    Else
        grid1.Col = 8
    End If
ElseIf Col = 8 Then
    If Row < grid1.Rows - 1 Then
        grid1.Row = Row + 1
        grid1.Col = 0
        grid1.ShowCell grid1.Row, 0
    End If
End If
End Sub
Private Sub XTQ1_GotFocus()
myGotFocus XTQ1
End Sub
Private Sub XTQ1_LostFocus()
myLostFocus XTQ1
End Sub
Private Sub XTQ2_GotFocus()
myGotFocus XTQ2
End Sub
Private Sub XTQ2_LostFocus()
myLostFocus XTQ2
End Sub
Private Sub xDescA_GotFocus()
myGotFocus XDESCA
End Sub
Private Sub xDesca_LostFocus()
myLostFocus XDESCA
End Sub
Private Sub XSTORE_GotFocus()
myGotFocus xStore, False
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus xStore
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_No
End Sub
Private Sub xfilter_GotFocus()
myGotFocus xfilter
End Sub
Private Sub xfilter_LostFocus()
myLostFocus xfilter
End Sub
Function SeekBarcode(pRow, pItem) As String
    With grid1
        
        If pItem <> "" Then
            If GetDesca("SELECT ITEM FROM FILE1_10 WHERE ITEM = " & Val(pItem), con) = "" Then
                If UCase(Mid(pItem, 1, 1)) = "J" Then
                    cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE2 = " & MyParn(pItem), con)
                ElseIf UCase(pItem) = "S" Then
                    cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE2 = " & MyParn(pItem), con)
                ElseIf Len(pItem) >= 10 Then
                    cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE13 = " & MyParn(pItem) & " OR BARCODE = " & MyParn(pItem) & " OR BARCODE_GS1 = " & MyParn(pItem), con)
                End If
            Else
                cMyItem = pItem
            End If
            If cMyItem <> "" Then
                SeekBarcode = Val(cMyItem)
            End If
        End If
    End With
End Function

Sub UnPostStock(pDate, pstore)
Dim SUBStockTable As New ADODB.Recordset
SUBStockTable.Open "SELECT * FROM FILE0_10H WHERE STORE =" & MyParn(pstore) & " AND FILE0_10H.date = " & DateSq(pDate), con, adOpenKeyset, adLockReadOnly, adCmdText
With SUBStockTable
    Do While Not .EOF
        cString = " UPDATE FILE0_10H SET FILE0_10H.closed = 0 WHERE DOC_NO = " & MyParn(!DOC_NO)
        con.Execute cString
    
        If cBranch = "00" Then
            cString = " DELETE FROM FILE1_11 WHERE TYPE = 'Z' AND FILE1_11.DOC_ID = " & MyParn(!DOC_NO)
            con.Execute cString
        End If
        .MoveNext
    Loop
End With

End Sub
Sub ReBalStock(pDate, pstore)
Dim SUBStockTable As New ADODB.Recordset
Dim nBal As Double
cString = " DELETE FROM FILE0_10 WHERE FILE0_10.DOC_NO IN (SELECT DOC_NO FROM FILE0_10H WHERE STORE =" & MyParn(pstore) & "AND DATE = " & DateSq(pDate) & " ) "
con.Execute cString

SUBStockTable.Open "SELECT MAX(FILE0_100.DOC_NO) AS DOC_NO ,ITEM , SUM(RealBal) AS RealBal FROM FILE0_100 INNER JOIN FILE0_10H ON FILE0_100.DOC_NO = FILE0_10H.DOC_NO WHERE STORE = " & MyParn(pstore) & " AND FILE0_10H.DATE = " & DateSq(pDate) & " GROUP BY ITEM ", con, adOpenStatic, adLockReadOnly, adCmdText
With SUBStockTable
Do While Not .EOF
        
    nBal = RetItemBalance(!Item, pstore, pDate, con, xCode.BoundText)
    SubReplaceItem_2 !Item, !RealBal, nBal, !RealBal - nBal, !DOC_NO
   .MoveNext
Loop
End With
End Sub
Sub PostStock(pDate, pstore)
Dim SUBStockTable As New ADODB.Recordset
SUBStockTable.Open "SELECT * FROM FILE0_10H WHERE STORE = " & MyParn(pstore) & " AND FILE0_10H.date = " & DateSq(pDate), con, adOpenStatic, adLockReadOnly, adCmdText
With SUBStockTable
    Do While Not .EOF
        cString = " UPDATE FILE0_10H SET FILE0_10H.closed = 0 WHERE DOC_NO = " & MyParn(!DOC_NO)
        con.Execute cString
        
        cString = " UPDATE FILE0_10H SET ISNEW = 1 , ISCLOSED = 1 , FILE0_10H.closed = 1 WHERE DOC_NO = " & MyParn(!DOC_NO)
        con.Execute cString
        .MoveNext
    Loop
End With
End Sub
Private Sub SubReplaceItem_2(pItem, nReal, nComputer, nDiffer, pDoc_no)
Dim aInsert(4, 1)
With grid1
    aInsert(0, 0) = "doc_no"
    aInsert(0, 1) = addstring(pDoc_no)
        
    aInsert(1, 0) = "item"
    aInsert(1, 1) = addstring(pItem)
        
    aInsert(2, 0) = "RealBal"
    aInsert(2, 1) = nReal
        
    aInsert(3, 0) = "ComputerBal"
    aInsert(3, 1) = nComputer
        
    aInsert(4, 0) = "Differ"
    aInsert(4, 1) = nDiffer
    con.Execute CreateInsert(aInsert, "FILE0_10")
End With
End Sub
Sub SubDoStock()
Dim cStr1   As String

On Error Resume Next
cStr1 = " CREATE VIEW Q_SUBSTOCK AS SELECT     FILE0_100.item FROM FILE0_100 INNER JOIN FILE0_10H ON FILE0_10H.DOC_NO = FILE0_100.doc_no WHERE FILE0_10H.STORE = " & MyParn(xStore.BoundText) & " AND FILE0_10H.DATE >= " & DateSq(xDate1.text) & " AND FILE0_10H.DATE <= " & DateSq(xDate2.text) & " group by FILE0_100.item "
con.Execute " DROP VIEW Q_SUBSTOCK"
createCommand cStr1, con
Err.Clear
End Sub
Private Sub grid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grid1.EditMaxLength = IIf(Col = 8, 20, 0)
End Sub
Private Function openCardTable(Optional pDoc_no As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String
Set CardTable = New ADODB.Recordset

cString = "SELECT TOP 1 * from FILE0_10H WHERE DOC_NO IS NOT NULL "
If lIsBranchStore Then cString = cString & " AND BRANCH = " & MyParn(cBranch)
If pSign = "=" Then
    If pDoc_no <> "" Then cWhere = " AND DOC_NO  " & pSign & addstring(pDoc_no)
Else
    If pDoc_no <> "" Then cWhere = " AND DOC_NO  " & pSign & addstring(pDoc_no)
End If
If sDoc_no <> "" Then cFilter = cFilter & " and DOC_NO = " & MyParn(sDoc_no)

' «·«Œ Ì«—« 
If cFilter <> "" Then cWhere = cWhere & cFilter
If cWhere <> "" Then cString = cString & cWhere
If pSign = "<" Or pSign = "<=" Then
    cString = cString & " order by doc_no desc"
ElseIf pSign = ">=" Or pSign = ">" Then
    cString = cString & " order by DATE DESC , doc_no ASC"
End If

Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Function

Private Sub CmdNext_Click()
openCardTable xDoc_No.text, ">"
If CardTable.EOF Then openCardTable xDoc_No.text
myload
End Sub
Private Sub CmdPrevious_Click()
openCardTable xDoc_No.text, "<"
If CardTable.EOF Then openCardTable xDoc_No.text
myload
End Sub
Private Sub CmdFirst_Click()
openCardTable , ">"
If Not CardTable.EOF Then
    myload
Else
    myDefine
End If
End Sub
Private Sub CmdLast_Click()
openCardTable , "<"
If Not CardTable.EOF Then
    myload
Else
    myDefine
End If
End Sub

Private Sub cmd_ex_Click22()
    Dim xl As New Excel.Application
    Dim xlsheet As Excel.Worksheet
    Dim xlwbook As Excel.Workbook
    Dim cFileName As String
    Dim nCount As Double, cItem As String, nQTY As Double
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
'           myreplace
            For nRow = 1 To nCount
                If xlsheet.Cells(nRow, 1) <> "" Then
                    cItem = (xlsheet.Cells(nRow, 1))
                    nQTY = Val((xlsheet.Cells(nRow, 2)) & "")
                    cItem = SeekBarcode(grid1.Rows - 1, cItem)
                    If cItem <> "" Then
                        .AddItem ""
                        grid1.TextMatrix(grid1.Rows - 1, 0) = cItem
                        grid1.TextMatrix(grid1.Rows - 1, 8) = nQTY
                        Dim ItemTable As New ADODB.Recordset
                        Set ItemTable = ItemFind(cItem, con)
                        If Not (ItemTable.EOF And ItemTable.BOF) Then
                            grid1.TextMatrix(grid1.Rows - 1, 1) = ItemTable!MOSM
                            grid1.TextMatrix(grid1.Rows - 1, 2) = GetDesca("select desca from fact where code = " & MyParn(ItemTable!Fact), con)
                            grid1.TextMatrix(grid1.Rows - 1, 3) = ItemTable!SUPP & ""
                            grid1.TextMatrix(grid1.Rows - 1, 4) = ItemTable!modelfact0
                            grid1.TextMatrix(grid1.Rows - 1, 5) = ItemTable!DESCA
                            grid1.TextMatrix(grid1.Rows - 1, 6) = ItemTable!color
                            grid1.TextMatrix(grid1.Rows - 1, 7) = ItemTable!scal
                        Else
                            GRID2.AddItem ""
                            GRID2.TextMatrix(GRID2.Rows - 1, 0) = (xlsheet.Cells(nRow, 1))
                            GRID2.TextMatrix(GRID2.Rows - 1, 1) = nQTY
                        End If
                    Else
                        GRID2.AddItem ""
                        GRID2.TextMatrix(GRID2.Rows - 1, 0) = (xlsheet.Cells(nRow, 1))
                        GRID2.TextMatrix(GRID2.Rows - 1, 1) = nQTY
                    End If
                End If
            Next nRow
            End With
            MsgBox "  „ «·≈÷«›… "
        End If
    End If
End Sub

