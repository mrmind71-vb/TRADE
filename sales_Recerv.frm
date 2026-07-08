VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form sales_Recerv 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10095
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10095
   ScaleWidth      =   15375
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CheckBox XISRET 
      Alignment       =   1  'Right Justify
      Caption         =   "„— Ã⁄"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11565
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   9090
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CommandButton cmd_ex 
      BackColor       =   &H00FFFFC0&
      Caption         =   " «” Ì—«œ „‰ «·«ﬂ”Ì· "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   13050
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   9045
      Width           =   2175
   End
   Begin VB.Frame Frame8 
      Height          =   780
      Left            =   5895
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   -45
      Width           =   7035
      Begin VB.OptionButton TO_EXEL 
         Alignment       =   1  'Right Justify
         Caption         =   "√’‰«› „Œ ·›… ⁄‰ «·≈” ·«„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   135
         Width           =   1455
      End
      Begin VB.OptionButton TO_EXEL 
         Alignment       =   1  'Right Justify
         Caption         =   "«’‰«› «ﬂ»— «·ÿ·»Ì… "
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
         Index           =   2
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   270
         Width           =   1320
      End
      Begin VB.OptionButton TO_EXEL 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ· «·«’‰«›"
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
         Index           =   0
         Left            =   3105
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   270
         Value           =   -1  'True
         Width           =   1230
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   555
         Left            =   4410
         TabIndex        =   47
         Top             =   135
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   979
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
         Picture         =   "sales_Recerv.frx":0000
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   12960
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   -45
      Width           =   5910
      Begin VB.CommandButton CmdDel 
         BackColor       =   &H008080FF&
         Caption         =   "Õ–› «·„” ‰œ"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1530
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton cmdNewInv 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   3000
         MaskColor       =   &H00FFFFFF&
         Picture         =   "sales_Recerv.frx":27FB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1410
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4455
         Picture         =   "sales_Recerv.frx":4DA7
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "sales_Recerv.frx":757A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Height          =   870
      Left            =   3555
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   8955
      Width           =   4155
      Begin VB.Label XTIME2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   495
         Width           =   2220
      End
      Begin VB.Label xUserName2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   495
         Width           =   1815
      End
      Begin VB.Label XTIME1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   135
         Width           =   2220
      End
      Begin VB.Label xUserName1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   135
         Width           =   1815
      End
      Begin VB.Label xusercode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   -270
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.Frame Frame7 
      Height          =   1860
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   -90
      Visible         =   0   'False
      Width           =   2265
      Begin VB.CheckBox xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·„” ‰œ „€·ﬁ"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   405
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1890
         Width           =   1320
      End
      Begin VB.Frame FRM_CLOSED 
         Height          =   1725
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   135
         Visible         =   0   'False
         Width           =   2040
         Begin Threed.SSCommand cmd_closed 
            CausesValidation=   0   'False
            Height          =   420
            Left            =   90
            TabIndex        =   23
            Top             =   1215
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   741
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "≈€·«ﬁ «·„” ‰œ"
         End
         Begin Threed.SSCommand cmd_CLOSEDDATE 
            CausesValidation=   0   'False
            Height          =   1005
            Left            =   1035
            TabIndex        =   24
            Top             =   180
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   1773
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
            Picture         =   "sales_Recerv.frx":99E6
            Caption         =   "≈€·«ﬁ › —…"
            Alignment       =   8
            PictureAlignment=   6
         End
         Begin Threed.SSCommand cmd_open 
            CausesValidation=   0   'False
            Height          =   1005
            Left            =   90
            TabIndex        =   25
            Top             =   180
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   1773
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
            Picture         =   "sales_Recerv.frx":C4BB
            Caption         =   "› Õ › —…"
            Alignment       =   8
            PictureAlignment=   6
         End
      End
      Begin VB.Label xF_date 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1890
         Width           =   1995
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "«·—’Ìœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4455
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   45
      Visible         =   0   'False
      Width           =   1455
      Begin VB.Label xBalance 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   180
         Width           =   780
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   4455
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   720
      Width           =   1455
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
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "sales_Recerv.frx":EFED
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "sales_Recerv.frx":11350
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   630
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5940
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   12885
      Begin VB.TextBox xDoc_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10575
         Locked          =   -1  'True
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   135
         Width           =   1185
      End
      Begin VB.TextBox XREC_NO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3015
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   135
         Width           =   1365
      End
      Begin VB.TextBox xinv_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5400
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   135
         Width           =   1365
      End
      Begin VB.TextBox xpo_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7875
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   135
         Width           =   1590
      End
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2970
         MaxLength       =   75
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   3570
      End
      Begin VB.TextBox xDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   810
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   135
         Width           =   1230
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   360
         Left            =   90
         TabIndex        =   1
         Top             =   540
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo XSTORE2 
         Height          =   330
         Left            =   3015
         TabIndex        =   51
         Top             =   585
         Width           =   3750
         _ExtentX        =   6615
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
      Begin MSDataListLib.DataCombo xcode 
         Height          =   330
         Left            =   7875
         TabIndex        =   59
         Top             =   585
         Width           =   3885
         _ExtentX        =   6853
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
      Begin VB.Label Label3 
         Caption         =   "«·⁄„Ì· :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11880
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   540
         Width           =   750
      End
      Begin VB.Label Label9 
         Caption         =   "≈–‰  ”·Ì„ "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   180
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„Œ“‰  ”·Ì„ "
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
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   585
         Width           =   930
      End
      Begin VB.Label Label4 
         Caption         =   "—ﬁ„ ›« Ê—… ÷—Ì»Ì… :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6795
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   180
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11835
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   180
         Width           =   885
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·ÿ·»Ì… :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9495
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   1020
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
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   -270
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label xtime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   135
         Width           =   690
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ«  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1125
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2130
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·„Œ“‰ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2130
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   705
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -3240
      Top             =   6660
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
      Left            =   -855
      Top             =   6885
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   -3060
      Top             =   6975
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
      Left            =   -1755
      Top             =   6885
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
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   11
      Top             =   9945
      Visible         =   0   'False
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7080
      Left            =   135
      TabIndex        =   3
      Top             =   1890
      Width           =   18735
      _cx             =   33046
      _cy             =   12488
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   630
      Picture         =   "sales_Recerv.frx":138C9
      RightToLeft     =   -1  'True
      ScaleHeight     =   2400
      ScaleWidth      =   3165
      TabIndex        =   13
      Top             =   3150
      Visible         =   0   'False
      Width           =   3165
   End
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   6480
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
   Begin VB.Frame Frame6 
      Height          =   645
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   9000
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   29
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
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
         Picture         =   "sales_Recerv.frx":16FA0
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "sales_Recerv.frx":19170
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
         TabIndex        =   30
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
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
         Picture         =   "sales_Recerv.frx":1B2B8
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "sales_Recerv.frx":1D480
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
         TabIndex        =   31
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
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
         Picture         =   "sales_Recerv.frx":1F5CF
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "sales_Recerv.frx":217AF
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
         TabIndex        =   32
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
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
         Picture         =   "sales_Recerv.frx":2390A
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "sales_Recerv.frx":25AC6
      End
   End
   Begin Threed.SSCommand CMD_PRINT 
      Height          =   465
      Left            =   2430
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   820
      _Version        =   196610
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "sales_Recerv.frx":27C15
      Caption         =   "ÿ»«⁄… «·›« Ê—…  "
      Alignment       =   1
      PictureAlignment=   3
   End
   Begin MSAdodcLib.Adodc data12 
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
   Begin Threed.SSCommand CMD_PRINT2 
      Height          =   645
      Left            =   2430
      TabIndex        =   53
      Top             =   315
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   1138
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
      Caption         =   "ÿ»«⁄… √’‰«› „Œ ·›… ⁄‰ «·≈” ·«„ Ê «·ÿ·»Ì… "
      Alignment       =   1
      ButtonStyle     =   2
      PictureAlignment=   3
   End
   Begin Threed.SSCommand cmdBarCode 
      Height          =   555
      Left            =   2385
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   990
      Width           =   1860
      _ExtentX        =   3281
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
      Picture         =   "sales_Recerv.frx":2A024
      Caption         =   " ÕÊÌ· ··»«—ﬂÊœ"
      Alignment       =   1
      PictureAlignment=   3
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
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
Attribute VB_Name = "sales_Recerv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temptable As New ADODB.Recordset
Dim lManSales As Boolean
Public sDoc_no As String, lSave As Boolean
Dim cFilter As String
Dim tCard As ADODB.Recordset
Dim nLeftMargin, nCardHeight, nPageWidth, nRightMargin
Dim nRound As Long, clist1 As String, cList2 As String
Dim CardTable As ADODB.Recordset
Public bRetvalue As Boolean
Dim cDefBox As String, cDefClient As String, cDefClientDesca As String, cDefCasher As String, cDefStore As String, cdefman As String
Dim oSearchItem As New Search3, oSearchDoc   As New Search3, oSearchCl1 As New Search3, bMarket As Boolean
Dim oSearchCl2 As New Search3, oSearchCl3 As New Search3
Dim bedit As Boolean
Dim cFile As String, cCodeDesca As String
Dim con As New ADODB.Connection
Dim formMode, dDateLast As String
Public myPublic As Integer
Const LoadMode = 0, DefineMode = 1
Sub ItemsLookup()
    ItemsLookupAll Me, oSearchItem
End Sub
Private Function myreplace(Optional nRow As Long = -1, Optional bModel As Boolean) As Boolean
Dim aInsert(10, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.text)


aInsert(1, 0) = "[Date]"
aInsert(1, 1) = addDate(xDate.text)

aInsert(2, 0) = "store"
aInsert(2, 1) = addstring(xStore.BoundText)

aInsert(3, 0) = "[Notes]"
aInsert(3, 1) = addstring(xNotes.text)

aInsert(4, 0) = "PO_NO"
aInsert(4, 1) = addstring(xpo_no.text)

aInsert(5, 0) = "INV_NO"
aInsert(5, 1) = addstring(xinv_no.text)

aInsert(6, 0) = "STORE2"
aInsert(6, 1) = addstring(XSTORE2.BoundText)

aInsert(7, 0) = "REC_NO"
aInsert(7, 1) = addstring(XREC_NO.text)

aInsert(8, 0) = "CODE"
aInsert(8, 1) = addstring(xCode.BoundText)



'con.BeginTrans
'''On Error GoTo myerror
If xDoc_No.Tag = DefineMode Then
    xDoc_No.text = RetZero(IncRec(GetDesca("select MAX(doc_no) from file6_52h ", con)), 6)
    If xDoc_No.text = "000000" Then xDoc_No.text = "000001"
    
    aInsert(9, 0) = "userName1"
    aInsert(9, 1) = addstring(xUserName1.Caption)

    aInsert(10, 0) = "TIME1"
    aInsert(10, 1) = addDate("GETDATE()")

    aInsert(0, 1) = addstring(xDoc_No.text)
    con.Execute CreateInsert(aInsert, "file6_52H")
Else
    aInsert(9, 0) = "userName2"
    aInsert(9, 1) = addstring(xUserName2.Caption)

    aInsert(10, 0) = "TIME2"
    aInsert(10, 1) = addDate("GETDATE()")
    
    con.Execute CreateUpdate(aInsert, "file6_52H", " where doc_no = " & addstring(xDoc_No.text))
End If

If nRow > 0 Then myreplaceGrd nRow

'con.CommitTrans
myreplace = True
Exit Function
myerror:
prog1.Visible = False
MsgBox Err.Description
'con.RollbackTrans
Err.Clear
xDoc_No.Tag = 1
End Function
Sub myProc()
'''On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    Dim bNew As Boolean
'    bNew = grid1.Row = grid1.Rows - 1
    If ActiveControl.Name = grid1.Name Then
        grid1.TextMatrix(grid1.Row, 1) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    ElseIf ActiveControl.Name = LOOKPH.Name Then
        grid1.TextMatrix(grid1.Row, 1) = ItemLookPh.grid1.TextMatrix(ItemLookPh.grid1.Row, 0)
    ElseIf ActiveControl.Name = LOOKPH2.Name Then
        grid1.TextMatrix(grid1.Row, 1) = ItemLookPh2.grid1.TextMatrix(ItemLookPh2.grid1.Row, 0)
    End If
    GrdDesc grid1.TextMatrix(grid1.Row, 1), grid1.Row
    grid1_AfterEdit grid1.Row, grid1.Col
    If Not bNew Then
        If ActiveControl.Name = grid1.Name Then Unload oSearchItem Else Unload ItemLookPh
        CellPos 13, grid1.Row, grid1.Col
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    xDoc_No.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    xDate.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 2)
    Unload oSearchDoc
    myUndo
ElseIf ActiveControl.Name = xpo_no.Name Then
    xpo_no.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmd_closed_Click()
    If CardTable!ISCLOSED Then
        con.Execute " update file6_52h set isclosed = 0 where doc_no = " & MyParn(xDoc_No.text)
    Else
        con.Execute " update file6_52h set isclosed = 1 where doc_no = " & MyParn(xDoc_No.text)
    End If
    myUndo
End Sub
Private Sub cmd_CLOSEDDATE_Click()
Dim DDate1 As Date, DDate2 As Date
DDate1 = TurnValue(InputBox("", "≈€·«ﬁ „” ‰œ«  „‰  «—ÌŒ", "1-1"), "", "1-1-2000")
DDate2 = TurnValue(InputBox("", "≈€·«ﬁ „” ‰œ«  Õ Ï  «—ÌŒ", Format(Date, "DD-MM-YYYY")), "", "1-1-2000")
If Not IsDate(DDate2) And IsDate(DDate1) Then DDate2 = dDate
If IsDate(DDate1) And IsDate(DDate2) Then
    If MsgBox("”Ê› Ì „ ≈€·«ﬁ „” ‰œ«  «·› —…", vbOKCancel) = vbOK Then
        con.Execute " update file6_52h set  isclosed = 1 where DATE > = " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2), nRec
        MsgBox " „ ≈€·«ﬁ „” ‰œ«  «·› —…" & nRec
        myUndo
    End If
End If
End Sub


Private Sub cmd_excel_Click()
    With grid1
    If TO_EXEL(0).Value <> 0 Then
        ToFileExel2 grid1, , , , , 1.1, , , , , , Me
    End If

    If TO_EXEL(2).Value <> 0 Then
        For nRow = 1 To .Rows - 1
            If .ValueMatrix(nRow, 9) > .ValueMatrix(nRow, 14) Then
                .RowHidden(nRow) = True
            End If
        Next nRow
        ToFileExel2 grid1, , , , , 1.1, , , , , , Me
        Fixgrd
    End If

    If TO_EXEL(1).Value <> 0 Then
        For nRow = 1 To .Rows - 1
            If .ValueMatrix(nRow, 9) <> .ValueMatrix(nRow, 15) Then
                .RowHidden(nRow) = True
            End If
        Next nRow
        ToFileExel2 grid1, , , , , 1.1, , , , , , Me
        Fixgrd
    End If
    End With
End Sub

Private Sub cmd_open_Click()
    Dim DDate1 As Date, DDate2 As Date
    DDate1 = TurnValue(InputBox("", "› Õ  „” ‰œ«  „‰  «—ÌŒ", ""), "", "1-1-2000")
    DDate2 = TurnValue(InputBox("", "› Õ „” ‰œ«  Õ Ï  «—ÌŒ", ""), "", "1-1-2000")
    If Not IsDate(DDate2) And IsDate(DDate1) Then DDate2 = dDate
    If IsDate(DDate1) And IsDate(DDate2) Then
        If MsgBox("”Ê› Ì „ › Õ „” ‰œ«  «·› —…", vbOKCancel) = vbOK Then
            con.Execute " update file6_52h set isclosed = 0 where DATE > = " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2), nRec
            MsgBox " „ › Õ „” ‰œ«  «·› —…" & nRec
            myUndo
        End If
    End If
End Sub
Private Sub CMD_PRINT_Click()
    xTax.text = Round(Val(xTotDisc.Caption) * ((Val(xRTax.text)) / 100), 2)
    xTotal.text = Round(nTotalItem - Val(xDiscount.text) + Val(xTax.text), 2)
    con.Execute " UPDATE file6_52H SET TAX = " & Val(xTax.text) & " WHERE DOC_NO = " & MyParn(xDoc_No.text)
    myUndo
    If xinv_no.text = "" Then MsgBox "  ”ÃÌ· —ﬁ„ «·›« Ê—…"
    If xpo_no.text = "" Then MsgBox "  ÕœÌœ —ﬁ„ «·ÿ·»Ì… "
    
    If xinv_no.text = "" Or xpo_no.text = "" Then Exit Sub
'    doprint 0
End Sub
Private Sub cmd_tsal_Click()
    TDaySal.Show
End Sub

Private Sub cmdBarCode_Click()
    addBarCode
    BarCodeFRM.Show 1

End Sub

Private Sub CmdDel_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    
    con.Execute "Delete  From file6_52 where Doc_No = " & MyParn(xDoc_No.text)
    con.Execute "Delete  From file6_52H where Doc_No = " & MyParn(xDoc_No.text)
    openCardTable
    CmdNewInv_Click
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdInform_Click()
    CardLookup
End Sub
Private Sub CmdNewInv_Click()
    myDefine
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
mysave
'on.Execute " UPDATE       FILE6_51H SET    ISPOSTED = 1  FROM  file6_52H INNER JOIN FILE6_51H ON file6_52H.PO_NO = FILE6_51H.DOC_NO"
'On Error Resume Next
Err.Clear
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdUndo_Click()
    myUndo
    grid1.ShowCell grid1.Rows - 1, 1
    grid1.Select grid1.Rows - 1, 1
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
End Sub
Private Sub Form_Activate()
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If xClosed.Value = 1 Then Exit Sub

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If

End Sub
Private Sub Form_Load()
Dim ImgXTmp As ImgX
Set ImgXTmp = New ImgX

bedit = True
openCon con
Set ImgXTmp = Nothing
Me.Caption = dSalesDate
nRound = 2

FRM_CLOSED.Visible = bopt2
Frame7.Visible = bopt2
Dim aRet As Variant
'aRet = aGetDesca("select code,desca from file3_10 where CODE = '0000' ")

Me.Caption = " ≈” ·«„ ÿ·»Ì… «„«“Ê‰"

Set data1.Recordset = myRecordSet("SELECT * FROM FILE0_40", con)
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

Set data2.Recordset = myRecordSet("SELECT * FROM FILE0_52", con)
Set XSTORE2.RowSource = data2
XSTORE2.ListField = "Desca"
XSTORE2.BoundColumn = "Code"


Set data2.Recordset = myRecordSet("SELECT * FROM FILE3_10", con)
Set xCode.RowSource = data2
xCode.ListField = "Desca"
xCode.BoundColumn = "Code"

Set grid1.DataSource = DATA11

myUndo


End Sub
Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next

CardTable.Close
Set CardTable = Nothing

closeCon con
Unload Search3
Unload oSearchDoc

Set SalesFrm = Nothing
Err.Clear
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
Dim bNew As Boolean

With grid1
If Not validRow(Row) Then Exit Sub
If myreplace(Row) Then
    HandleCntEdit
    bNew = grid1.TextMatrix(Row, .Cols - 1) = ""
End If
myloadgrd
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub grid1_EnterCell()
If (xClosed.Value <> 0) Then
    grid1.Editable = flexEDNone
    Exit Sub
End If
With grid1
    If .Col = 1 Or .Col = 8 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub Grid1_GotFocus()
On Error Resume Next
If xpo_no.text = "" Then xpo_no.SetFocus
If grid1.Rows < 2 Then Exit Sub
If grid1.Row = 0 Then
    grid1.Row = 1
    grid1.Col = 1
End If
grid1_EnterCell
Err.Clear
Exit Sub
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And xClosed.Value = 0 Then
    ItemsLookupAll Me, oSearchItem
End If
End Sub
Private Sub Grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    With grid1
    If grid1.Row = grid1.Rows - 1 And grid1.Row > 1 And Trim(grid1.TextMatrix(grid1.Row, 1)) = "" Then
        KeyAscii = 0
        If xClosed.Value = 0 Then cmdSave_Click
        Err.Clear
    End If
    End With
End If
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim cId As String, cMyItem As Double
With grid1
    If Col = 1 Then
        If .EditText <> "" And Len(.EditText) > 2 Then
'           cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE_GS1 = " & MyParn(.EditText) & " or barcode13 =  " & MyParn(.EditText) & " or item = " & Val(.EditText))
            cMyItem = Item_BARCODE(.EditText, con)
            If cMyItem <> 0 Then
                cItem = GetDesca("SELECT ITEM FROM  file6_51 inner join file6_51h on file6_51.doc_no = file6_51h.doc_no WHERE file6_51h.doc_no = " & MyParn(xpo_no.text) & " AND ITEM = " & (cMyItem), con)
                If cItem = cMyItem Then
                    .EditText = cMyItem
                    GrdDesc .EditText, Row
                Else
                    n = Beep(2000, 1000)
                    MsgBox " ﬂÊœ €Ì— „ÿ·Ê» "
                    Cancel = True
                End If
            Else
                n = Beep(2000, 500)
                MsgBox " »«—ﬂÊœ €Ì— ’ÕÌÕ "
            End If
        End If
    End If
End With
End Sub
Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 And xClosed.Value = 0 Then CLIENTLOOKUP
End Sub
Private Sub xDate_Validate(Cancel As Boolean)
If Not IsDate(xDate.text) Then Cancel = True
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Not IsDate(xDate.text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xStore.BoundText = "" Then
    If Not bIgMsg Then MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
    Exit Function
End If
If xpo_no.text = "" Then
    If Not bIgMsg Then MsgBox " ”ÃÌ· —ﬁ„ ÿ·»Ì… «„«“Ê‰ "
    Exit Function
End If
MYVALID = True
End Function
Private Sub myload(Optional bLeaveBal As Boolean = False)
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
xDoc_No.text = CardTable!doc_no
xinv_no.text = CardTable!INV_NO & ""
xpo_no.text = CardTable!po_no & ""
XREC_NO.text = CardTable!REC_NO & ""

xCode.BoundText = CardTable!CODE & ""

xDate.text = Format(CardTable!Date, "dd-mm-yyyy")
xStore.BoundText = CardTable!STORE & ""
XSTORE2.BoundText = CardTable!STORE2 & ""

xNotes.text = CardTable!NOTES & ""
xtime.Caption = Format(CardTable!Time, "hh:nn")
myloadgrd
Handlecontrols LoadMode
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myDefine()
cmd_closed.BackColor = &H8080FF
cmd_closed.Caption = "› Õ „” ‰œ"
xClosed.Visible = False
xClosed.Value = ssCBUnchecked

xDoc_No.text = RetZero(IncRec(GetDesca("select MAX(doc_no) from file6_52h ", con)), 6)
If xDoc_No.text = "000000" Then xDoc_No.text = "000001"
xpo_no.text = ""
XREC_NO.text = ""
xCode.BoundText = ""
XISRET.Value = 1
xDoc_No.Tag = DefineMode

xUserName2.Caption = ""
XTIME1.Caption = ""
XTIME2.Caption = ""
xDate.text = Format(Date, "DD-MM-YYYY")
xStore.BoundText = "1"
XSTORE2.BoundText = ""

xNotes.text = ""
xtime.Caption = Format(Time, "hh:nn")
grid1.Rows = 1
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
xDoc_No.Tag = 0

Fixgrd
Handlecontrols DefineMode
cmdSave.Enabled = True
CmdDel.Enabled = True
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = nMode = LoadMode
cmdSave.Enabled = (grid1.Rows > 2) And xClosed.Value = 0
CmdDel.Enabled = xClosed.Value = 0
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)

xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode

End Sub
Private Function retBool(cFieldName) As Boolean
If Not (CardTable.EOF Or CardTable.BOF) Then
    retBool = CardTable(cFieldName)
End If
End Function
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload True
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If xClosed.Value = 1 Then Exit Sub
If KeyCode = 46 And grid1.Row = grid1.Rows - 1 And xClosed.Value = 0 And grid1.Rows > 3 Then
    grid1.Select grid1.Rows - 2, 0
End If
If KeyCode = 116 Then
    cmdSave_Click
End If
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And CmdDel.Enabled And grid1.Rows > 3 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        RemoveItem (grid1.Row)
        grid1.Select grid1.Rows - 1, 1
        grid1.ShowCell grid1.Rows - 1, 1
        
        
        'UpdateHeader
        MakeSerial grid1.Row
    End If
End If
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
End Sub
Private Function GrdDesc(sitem As String, Row As Long) As Boolean
Dim lMosmDisc As Boolean
Dim cMyItem As String
With grid1
If sitem = "" Then Exit Function
If IsNumeric(sitem) Then
    If Len(sitem) > 2 Then
        Dim loctable As ADODB.Recordset
        If Len(sitem) > 10 Then
            MsgBox "»«—ﬂÊœ €Ì— ’ÕÌÕ"
            Exit Function
        End If
        Set loctable = ItemFind(sitem, con)
        If Not (loctable.EOF And loctable.BOF) Then
            grid1.TextMatrix(Row, 2) = loctable!MOSM
            grid1.TextMatrix(Row, 3) = GetDesca("select desca from fact where code = " & MyParn(loctable!Fact), con)
            grid1.TextMatrix(Row, 5) = loctable!modelfact0 & ""
            grid1.TextMatrix(Row, 5) = loctable!DESCA
            grid1.TextMatrix(Row, 6) = loctable!scal
            grid1.TextMatrix(Row, 7) = loctable!color
            grid1.TextMatrix(Row, 8) = 1
            grid1.TextMatrix(Row, 9) = loctable!BARCODE_GS1 & ""
            GrdDesc = True
        End If
    End If
End If
End With
End Function
Private Sub CardLookup(Optional pFilter As String = "")
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(6, 1)
Set Generalarray(0) = Me
'                                   0               1                   2                   3                                   4
Generalarray(1) = "SELECT   file6_52H.DOC_NO, file6_52H.PO_NO ,  FILE3_10.DESCA  ,file6_52H.INV_NO,file6_52H.date, CONVERT(VARCHAR(10), file6_52H.date, 111) , REC_NO " & _
                  " FROM    file6_52H LEFT JOIN FILE3_10 ON FILE3_10.CODE = FILE6_52H.CODE "
Generalarray(2) = "Order by doc_no  DESC "
Generalarray(3) = 6000
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„ -  «—ÌŒ"
listarray(0, 1) = "(%%INV_NO%% or %%PO_NO%% or %%REC_NO%% OR ##date## )"
   

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1200

GrdArray(1, 0) = "—ﬁ„ PO "
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«·⁄„Ì·"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "—ﬁ„ «·›« Ê—… "
GrdArray(3, 1) = 1500

GrdArray(4, 0) = "«· «—ÌŒ"
GrdArray(4, 1) = 0

GrdArray(5, 0) = "«· «—ÌŒ"
GrdArray(5, 1) = 1500

GrdArray(6, 0) = "—ﬁ„  ”·Ì„"
GrdArray(6, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.sId = "sales" & IIf(bopt3, "1", "0")
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub
Private Sub xDoc_No_Validate(Cancel As Boolean)
If xDoc_No.text = "" Then Cancel = True
End Sub
Private Sub xpo_no_LostFocus()
    If XSTORE2.BoundText = "" And xpo_no.text <> "" Then XSTORE2.BoundText = GetDesca("SELECT STORE FROM FILE6_51H WHERE DOC_NO = " & MyParn(xpo_no.text), con)
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
Private Sub Fixgrd()
With grid1

'                 0        1         2       3              4                   5               6           7         8          9
.FormatString = "„.|" & "CODE|" & "MOSM|" & "BRAND|" & "productId|" & "productDescription|" & "COLOR|" & "SIZE|" & "QTY|" & "BARCODE|"

.RowHeight(0) = 800
.WordWrap = True
.ColHidden(.Cols - 1) = True
.ColWidth(0) = 700
.ColWidth(1) = 2000
.ColWidth(2) = 1000
.ColWidth(3) = 4000
.ColWidth(4) = 1500
.ColWidth(5) = 4000
.ColWidth(6) = 1500
.ColWidth(7) = 1000
.ColWidth(8) = 1000
.ColWidth(9) = 2000

.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
.Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter

For nRow = 1 To .Rows - 1
    .RowHidden(nRow) = False
    .TextMatrix(nRow, 0) = nRow
Next nRow
.ShowCell .Rows - 1, 0
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
oSearchCl1.Caption = "«” ⁄·«„"
oSearchCl1.Show 1
End Sub

Private Sub XSTORE_GotFocus()
    myGotFocus xStore
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus xStore
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub xStore_Validate(Cancel As Boolean)
    If Trim(xStore.BoundText) = "" Then Cancel = True
End Sub
Private Function RemoveItem(nRow) As Boolean
''On Error GoTo myerror
con.BeginTrans
If grid1.TextMatrix(nRow, grid1.Cols - 1) <> "" Then
    con.Execute "Delete  From file6_52 where id = " & grid1.TextMatrix(nRow, grid1.Cols - 1)
End If
grid1.RemoveItem nRow
con.CommitTrans
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
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
    itemPrice = GetDesca("select PRICE from file1_10 where item = " & MyParn(cItem), con)
End Function
Private Function myreplaceGrd(nRow) As Boolean
On Error GoTo myerror
Dim aInsert(3, 1), nTQ As Double
With grid1
    For i = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.text)
        
        aInsert(1, 0) = "item"
        aInsert(1, 1) = addstring(grid1.TextMatrix(i, 1))
        
        aInsert(2, 0) = "quant"
        aInsert(2, 1) = Val(.TextMatrix(i, 8))

        aInsert(3, 0) = "row"
        aInsert(3, 1) = i
        
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, "file6_52")
        Else
            con.Execute CreateUpdate(aInsert, "file6_52", " where ID = " & grid1.TextMatrix(i, .Cols - 1))
    End If
        If nRow > 0 Then
            nTQ = Val(GetDesca("SELECT SUM(TREC-TORDER) FROM  Q_ORDER_AMAZON WHERE doc_no = " & MyParn(xpo_no.text) & " AND ITEM = " & Val(.TextMatrix(nRow, 1)), con) & "")
            If nTQ > 0 Then
                MsgBox " ≈” ·«„ ﬂ„Ì… «ﬂ»— „‰ «·„ÿ·Ê» "
                .Select nRow, 1
                .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbRed
            End If
        End If
'        If nRow <> 0 Then grid1.Row = Row
    Next
End With
myreplaceGrd = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Sub myloadgrd()
With grid1
    Dim cString As String, cFiled1 As String, cFiled2 As String
    '                 0        1         2       3          4                   5                   6           7         8          9
    .FormatString = "„.|" & "CODE|" & "MOSM|" & "BRAND|" & "productId|" & "productDescription|" & "COLOR|" & "SIZE|" & "QTY|" & "BARCODE GS1"
'                       0                   2               3               4               5               6               7                   8           9           10
     cString = "SELECT file6_52.ITEM, FILE1_10.MOSM, FACT.DESCA , FILE1_10.MODELFACT0, FILE1_10.DESCA, FILE1_10.COLOR, FILE1_10.SCAL  , file6_52.Quant, BARCODE_GS1 , file6_52.ID " & _
          " FROM file6_52 INNER JOIN FILE1_10 ON file6_52.ITEM = FILE1_10.ITEM INNER JOIN FACT ON file1_10.FACT = FACT.CODE"
    cString = cString & " WHERE DOC_NO = " & MyParn(xDoc_No.text)
    cString = cString & " ORDER BY file6_52.ID"
    Set DATA11.Recordset = myRecordSet(cString, con)
    .AddItem ""
End With
Fixgrd
grid1.ShowCell grid1.Rows - 1, 1
End Sub
Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_No
End Sub
Private Sub xdate_GotFocus()
myGotFocus xDate
End Sub
Private Function mysave(Optional bEnd As Boolean = True, Optional bPrint As Boolean = True) As Boolean
If Not MYVALID Then Exit Function
If Not myreplace Then Exit Function
If bEnd And lSave Then
    Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
    myDefine
Else
    myUndo
End If
End Function
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = True) As Boolean
With grid1
If Not MYVALID(bIgMsg) Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 1)) Then Exit Function
End With
validRow = True
End Function
Private Sub HandleCntEdit()
xDoc_No.Tag = LoadMode
xDoc_No.Enabled = False
CmdDel.Enabled = True
cmdSave.Enabled = xClosed.Value = 0 And grid1.Rows > 2

End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
KeyCode = 0
If Col = 1 And False Then
    grid1.Col = 8
    grid1.Row = Row
    grid1.ShowCell Row, 8
    grid1.Select Row, 8
    Exit Sub
ElseIf Col = 1 Then
    grid1.Row = Row + 1
    grid1.Col = 1
    grid1.ShowCell grid1.Row, 1
    Exit Sub
ElseIf Col >= 8 Then
    grid1.Row = Row + 1
    grid1.Col = 1
    grid1.ShowCell grid1.Row, 1
    Exit Sub
End If
Err.Clear
End Sub

Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
End Sub
Private Sub doprint_doc()
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim lPrice As Boolean
Dim nFBal As Double
Dim nPay As Double
Dim nBal As Double
Dim cAddres As String, cTaxNo As String

cStr1 = "SELECT file6_52.ITEM , FILE1_10.DESCA , FILE1_10.SCAL , FILE1_10.C_SCAL , FILE1_10.COLOR , file6_52.PRICE , SUM(file6_52.QUANT ) AS QUANT , FILE1_10.MODELFACT0 , SUM(file6_52.TOTAL) AS TOTAL  FROM FILE1_10 INNER JOIN file6_52 ON file6_52.ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(xDoc_No.text) & " GROUP BY file6_52.ITEM , FILE1_10.DESCA , FILE1_10.SCAL , FILE1_10.C_SCAL , FILE1_10.COLOR , file6_52.PRICE , FILE1_10.MODELFACT0"
sourcetable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
cAddres = GetDesca("SELECT ADDRESS FROM FILE3_10 WHERE CODE = " & MyParn(xCode.BoundText), con)
cTaxNo = GetDesca("SELECT EMAIL FROM FILE3_10 WHERE CODE = " & MyParn(xCode.BoundText), con)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
With sourcetable
Do While Not .EOF
    temptable.AddNew
    temptable!str1 = xDoc_No.text
    temptable!str2 = IIf(xCodeDesca2.Caption = "", xCodeDesca.Caption, xCodeDesca2.Caption)
    temptable!Date1 = xDate.text
    temptable!str6 = TurnValue(cAddres)
    temptable!str15 = TurnValue(cTaxNo)
    temptable!str14 = !modelfact0
    temptable!val3 = !TOTAL
    temptable!val2 = !price
    temptable!VAL1 = !Quant
    temptable!str8 = !DESCA
    temptable!str9 = !scal
    temptable!VAL19 = !C_SCAL
    temptable!STR7 = !color
    
    temptable!VAL7 = Val(xTotalItem.Caption)
    temptable!val9 = Val(xTotalQuant.Caption) - Val(xTotalQuant2.Caption)
    
    temptable!str18 = "≈Ã„«·Ï «’‰«›"
    temptable!val18 = Val(xTotalItem.Caption)
    temptable!val17 = 0
    temptable!STR19 = 0
    temptable!val15 = Val(xTax.text)
    temptable!val16 = Val(xTotal.text)
    temptable!str16 = MyOnly(Val(xTotal.text))
    temptable!str18 = "Œ’„ «·›« Ê—…"
    temptable!VAL7 = Val(xDiscount.text)
    temptable!Val8 = Val(xTotDisc.Caption)
    
    temptable.Update
    .MoveNext
Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\PRINT_INVOICE.rpt"
Main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub xPO_NO_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then OrderNo_LooK
End Sub
Sub OrderNo_LooK()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO,DATE, CONVERT(VARCHAR(10),[DATE],111),NOTES " & _
                  " FROM FILE6_51H WHERE doc_sales IS NULL AND ISPOSTED = 0 "
If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If

Generalarray(2) = "Order by Date , DOC_NO "
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ"
listarray(0, 1) = "(@@Doc_No@@6 OR " & _
                  " ##[DATE]##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1200

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·»Ì«‰"
GrdArray(3, 1) = 2000


searchArray = Array(Generalarray, listarray, GrdArray)
Load oSearchDoc
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub

Private Function addBarCode() As String
con.BeginTrans
'On Error GoTo myerror
con.Execute "DELETE FROM ADDPRINT WHERE BRANCH = " & MyParn(cBranch) & " AND DOC_NO = " & MyParn(xDoc_No.text)
cString = "INSERT INTO ADDPRINT(DOC_NO,ITEM,BRANCH,QUANT,ISPRINT)"

cString = cString & _
          " SELECT DOC_NO,ITEM," & addstring(cBranch) & ",QUANT,1 FROM FILE6_52"
cString = cString & turn(cString) & " DOC_NO = " & MyParn(xDoc_No.text)
con.Execute cString
 
con.CommitTrans
Inform " „ «· ÕÊÌ· »‰Ã«Õ"
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Sub cmd_ex_Click()
'    On Error GoTo myerror
    AddFrom_Order
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear

End Sub

Sub AddFrom_Order()
    Dim xl As New Excel.Application
    Dim xlsheet As Excel.Worksheet
    Dim xlwbook As Excel.Workbook
    Dim cDoc_No As String, nRow As Long
    Dim cBarCode As String
    Dim cModel As String
    Dim cItem As String, cENA As String
    Dim nQTY As Double, nRec As Double
    Dim nPrice As Double, CGROUP As String, cCodeGroup As String
    Dim nDisc As Double, cDesca2 As String
    Dim nNewItem As String, cCode As String, cSex As String
    Dim lAddTrue As Boolean, cModelAll As String
    Dim cStr1 As String, cStr2 As String, nCoutnR As Double, cDocNo As String, nRowDoc As Double
    prog1.Min = 0
    prog1.Value = 0
    prog1.Visible = True
    If MsgBox("≈÷«›… „— Ã⁄ «” ·«„ ÿ·»Ì… ⁄„Ì· „·› Excel", vbYesNo) = vbYes Then
        Common1.InitDir = ""
        Common1.FileName = ""
        Common1.Filter = "Excel (*.XLS*)|*.XLS*"
        Common1.ShowOpen
        cFileName = Common1.FileName
        If cFileName <> "" Then
            Set xlwbook = xl.Workbooks.Open(cFileName)
            Set xlsheet = xlwbook.Sheets.Item(1)
            
            For nRow = 2 To 10000
                If Len(xlsheet.Cells(nRow, 1)) > 1 Then
                    nRec = nRow
                End If
            Next nRow
            prog1.Max = nRec
            i = 0
            For nRow = 2 To nRec
                prog1.Visible = True
                prog1.Value = nRow
                If Len(xlsheet.Cells(nRow, 1)) = 0 Then Exit For
                nQTY = Val(xlsheet.Cells(nRow, 2))
                cItem = Val(xlsheet.Cells(nRow, 1))
                grid1.TextMatrix(nRow - 1, 0) = nRow - 1
                grid1.TextMatrix(nRow - 1, 1) = cItem
                GrdDesc cItem, nRow - 1
                If XISRET.Value = 0 Then
                    grid1.TextMatrix(nRow - 1, 8) = nQTY * -1
                Else
                    grid1.TextMatrix(nRow - 1, 8) = nQTY
                End If
                myreplaceGrd nRow - 1
                grid1.AddItem ""
            Next nRow
        End If
        MsgBox "  „ «÷«›… «·«’‰«› "
    End If
End Sub

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
Private Function openCardTable(Optional pDoc_No As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String
Set CardTable = New ADODB.Recordset
cFilter = ""
cFilterLook = ""
cString = "SELECT TOP 1 * from FILE6_52H"
If sDoc_no <> "" Then
    cFilter = cFilter & Tr(cFilter) & " DOC_NO = " & MyParn(sDoc_no)
End If
If pSign = "=" Then
    If pDoc_No <> "" Then cWhere = " DOC_NO  " & pSign & addstring(pDoc_No)
Else
    If pDoc_No <> "" Then cWhere = " DOC_NO  " & pSign & addstring(pDoc_No)
End If

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " Where " & cWhere
If pSign = "<" Or pSign = "<=" Then
    cString = cString & " order by doc_no desc"
ElseIf pSign = ">=" Or pSign = ">" Then
    cString = cString & " order by doc_no ASC"
End If
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Function

