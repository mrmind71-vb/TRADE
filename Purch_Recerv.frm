VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Purch_Recerv 
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10095
   ScaleWidth      =   15375
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      Height          =   780
      Left            =   5895
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   -45
      Visible         =   0   'False
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   270
         Value           =   -1  'True
         Width           =   1230
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   555
         Left            =   4410
         TabIndex        =   40
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
         Picture         =   "Purch_Recerv.frx":0000
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   12960
      RightToLeft     =   -1  'True
      TabIndex        =   34
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
         TabIndex        =   38
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
         Picture         =   "Purch_Recerv.frx":27FB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1410
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4455
         Picture         =   "Purch_Recerv.frx":4DA7
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Purch_Recerv.frx":757A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   35
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
      TabIndex        =   28
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   -270
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.Frame Frame7 
      Height          =   1860
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   16
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
         TabIndex        =   21
         Top             =   1890
         Width           =   1320
      End
      Begin VB.Frame FRM_CLOSED 
         Height          =   1725
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   135
         Visible         =   0   'False
         Width           =   2040
         Begin Threed.SSCommand cmd_closed 
            CausesValidation=   0   'False
            Height          =   420
            Left            =   90
            TabIndex        =   18
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
            TabIndex        =   19
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
            Picture         =   "Purch_Recerv.frx":99E6
            Caption         =   "≈€·«ﬁ › —…"
            Alignment       =   8
            PictureAlignment=   6
         End
         Begin Threed.SSCommand cmd_open 
            CausesValidation=   0   'False
            Height          =   1005
            Left            =   90
            TabIndex        =   20
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
            Picture         =   "Purch_Recerv.frx":C4BB
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
         TabIndex        =   22
         Top             =   1890
         Width           =   1995
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   2385
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
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
         Picture         =   "Purch_Recerv.frx":EFED
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "Purch_Recerv.frx":11350
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   630
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
   End
   Begin VB.Frame nhlj 
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
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   15000
      Begin VB.TextBox XINV_NO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   6345
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   180
         Width           =   1275
      End
      Begin VB.TextBox XNOTES 
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
         Left            =   90
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   585
         Width           =   5280
      End
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
         Left            =   12375
         Locked          =   -1  'True
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   180
         Width           =   1185
      End
      Begin VB.TextBox XORDER_NO 
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
         Left            =   9675
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   1410
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
         Left            =   3060
         TabIndex        =   1
         Top             =   180
         Width           =   2310
         _ExtentX        =   4075
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
      Begin MSDataListLib.DataCombo xcode 
         Height          =   330
         Left            =   9675
         TabIndex        =   45
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ „” ‰œ „‘ —Ì«  : "
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
         Left            =   7695
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   225
         Width           =   1845
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ«  "
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
         Left            =   5490
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   630
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "«·„Ê—œ : "
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
         Left            =   13725
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   540
         Width           =   750
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
         Left            =   13725
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   975
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
         Left            =   11160
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   225
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
         TabIndex        =   14
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
         TabIndex        =   4
         Top             =   135
         Width           =   690
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
         TabIndex        =   8
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
         Left            =   5460
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
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
      TabIndex        =   9
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
      TabIndex        =   2
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
      Picture         =   "Purch_Recerv.frx":138C9
      RightToLeft     =   -1  'True
      ScaleHeight     =   2400
      ScaleWidth      =   3165
      TabIndex        =   10
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
      TabIndex        =   23
      Top             =   9000
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   24
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
         Picture         =   "Purch_Recerv.frx":16FA0
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Purch_Recerv.frx":19170
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
         TabIndex        =   25
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
         Picture         =   "Purch_Recerv.frx":1B2B8
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Purch_Recerv.frx":1D480
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
         TabIndex        =   26
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
         Picture         =   "Purch_Recerv.frx":1F5CF
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Purch_Recerv.frx":217AF
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
         TabIndex        =   27
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
         Picture         =   "Purch_Recerv.frx":2390A
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Purch_Recerv.frx":25AC6
      End
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
   Begin Threed.SSCommand CMD_PRINT 
      Height          =   555
      Left            =   4005
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   45
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   979
      _Version        =   196610
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Purch_Recerv.frx":27C15
      Caption         =   "ÿ»«⁄… „” ‰œ «·ÿ·»Ì… "
      Alignment       =   1
      PictureAlignment=   3
   End
End
Attribute VB_Name = "Purch_Recerv"
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
Dim aInsert(7, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_no.text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = addDate(xDate.text)

aInsert(2, 0) = "store"
aInsert(2, 1) = addstring(xStore.BoundText)

aInsert(3, 0) = "[Notes]"
aInsert(3, 1) = addstring(XNOTES.text)

aInsert(4, 0) = "ORDER_NO"
aInsert(4, 1) = addstring(XORDER_NO.text)

aInsert(5, 0) = "CODE"
aInsert(5, 1) = addstring(xcode.BoundText)


'con.BeginTrans
'''On Error GoTo myerror
If xDoc_no.Tag = DefineMode Then
    xDoc_no.text = RetZero(IncRec(GetDesca("select MAX(doc_no) from file6_55h ", con)), 6)
    If xDoc_no.text = "000000" Then xDoc_no.text = "000001"
    
    aInsert(6, 0) = "userName1"
    aInsert(6, 1) = addstring(xUserName1.Caption)

    aInsert(7, 0) = "TIME1"
    aInsert(7, 1) = addDate("GETDATE()")

    aInsert(0, 1) = addstring(xDoc_no.text)
    con.Execute CreateInsert(aInsert, "file6_55H")
Else
    aInsert(6, 0) = "userName2"
    aInsert(6, 1) = addstring(xUserName2.Caption)

    aInsert(7, 0) = "TIME2"
    aInsert(7, 1) = addDate("GETDATE()")
    
    con.Execute CreateUpdate(aInsert, "file6_55H", " where doc_no = " & addstring(xDoc_no.text))
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
xDoc_no.Tag = 1
End Function
Sub myProc()
'''On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    Dim bNew As Boolean
    bNew = grid1.Row = grid1.Rows - 1
    If ActiveControl.Name = grid1.Name Then
        grid1.TextMatrix(grid1.Row, 1) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    End If
    GrdDesc grid1.TextMatrix(grid1.Row, 1), grid1.Row
    grid1_AfterEdit grid1.Row, grid1.Col
    grid1.Select grid1.Rows - 1, 1
ElseIf ActiveControl.Name = CmdInform.Name Then
    xDoc_no.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    xDate.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 2)
    Unload oSearchDoc
    myUndo
ElseIf ActiveControl.Name = XORDER_NO.Name Then
    XORDER_NO.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
ElseIf ActiveControl.Name = xcode.Name Then
    xcode.BoundText = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmd_closed_Click()
    If CardTable!ISCLOSED Then
        con.Execute " update file6_55h set isclosed = 0 where doc_no = " & MyParn(xDoc_no.text)
    Else
        con.Execute " update file6_55h set isclosed = 1 where doc_no = " & MyParn(xDoc_no.text)
    End If
    openCardTable
    myUndo
End Sub
Private Sub cmd_CLOSEDDATE_Click()
Dim DDate1 As Date, DDate2 As Date
DDate1 = TurnValue(InputBox("", "≈€·«ﬁ „” ‰œ«  „‰  «—ÌŒ", "1-1"), "", "1-1-2000")
DDate2 = TurnValue(InputBox("", "≈€·«ﬁ „” ‰œ«  Õ Ï  «—ÌŒ", Format(Date, "DD-MM-YYYY")), "", "1-1-2000")
If Not IsDate(DDate2) And IsDate(DDate1) Then DDate2 = dDate
If IsDate(DDate1) And IsDate(DDate2) Then
    If MsgBox("”Ê› Ì „ ≈€·«ﬁ „” ‰œ«  «·› —…", vbOKCancel) = vbOK Then
        con.Execute " update file6_55h set  isclosed = 1 where DATE > = " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2), nRec
        MsgBox " „ ≈€·«ﬁ „” ‰œ«  «·› —…" & nRec
        openCardTable
        myUndo
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
            con.Execute " update file6_55h set isclosed = 0 where DATE > = " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2), nRec
            MsgBox " „ › Õ „” ‰œ«  «·› —…" & nRec
            openCardTable
            myUndo
        End If
    End If
End Sub

Private Sub CMD_PRINT_Click()
    myUndo
    doprint

End Sub

Private Sub CmdDel_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    con.Execute "Delete  From file6_55 where Doc_No = " & MyParn(xDoc_no.text)
    con.Execute "Delete  From file6_55H where Doc_No = " & MyParn(xDoc_no.text)
    openCardTable
    CmdNewInv_Click
End If
Exit Sub
myerror:
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
    mydefine
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
End Sub
Private Sub cmdSave_Click()
    If Not MYVALID Then Exit Sub
    mysave
    Err.Clear
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub
Private Sub CmdUndo_Click()
    openCardTable
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

Me.Caption = " ≈” ·«„ ÿ·»Ì… ‘—«¡ „‰ „Ê—œ "

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FILE0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"


DATA3.ConnectionString = strCon
DATA3.RecordSource = "SELECT * FROM FILE4_10 ORDER BY DESCA "
Set xcode.RowSource = DATA3
xcode.ListField = "Desca"
xcode.BoundColumn = "Code"

Set grid1.DataSource = DATA11
DATA11.ConnectionString = strCon


openCardTable
If sDoc_no = "" Then
    CmdNewInv_Click
Else
    myUndo
End If

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
'''On Error GoTo myerror
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
If KeyCode = 114 And xClosed.Value = 0 Then
    ItemsLookupAll Me, oSearchItem
End If
If KeyCode = 112 And xClosed.Value = 0 Then
    ItemsLookupAll Me, oSearchItem, XORDER_NO.text
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
Dim cId As String, cMyItem As String
With grid1
    If Col = 1 Then
        If .EditText <> "" And Len(.EditText) > 2 Then
            cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE_GS1 = " & MyParn(.EditText) & " or barcode13 =  " & MyParn(.EditText) & " or item = " & Val(.EditText), con)
            If cMyItem <> "" Then
                cItem = GetDesca("SELECT ITEM FROM  file6_50 inner join file6_50h on file6_50.doc_no = file6_50h.doc_no WHERE file6_50h.doc_no = " & MyParn(XORDER_NO.text) & " AND ITEM = " & Val(cMyItem), con)
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
    If KeyCode = 112 And xClosed.Value = 0 Then SuppLookupAll Me, Search3
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
If XORDER_NO.text = "" Then
    If Not bIgMsg Then MsgBox " ”ÃÌ· —ﬁ„ ÿ·»Ì… «·„Ê—œ "
'    Exit Function
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
xDoc_no.text = CardTable!doc_no
XORDER_NO.text = CardTable!ORDER_NO & ""
XINV_NO.text = CardTable!INV_NO & ""
xcode.BoundText = CardTable!CODE & ""
xDate.text = Format(CardTable!Date, "dd-mm-yyyy")
xStore.BoundText = CardTable!STORE & ""
XNOTES.text = CardTable!NOTES & ""
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
xDoc_no.text = RetZero(IncRec(GetDesca("select MAX(doc_no) from file6_55h ", con)), 6)

If xDoc_no.text = "000000" Then xDoc_no.text = "000001"
XORDER_NO.text = ""
XINV_NO.text = ""
xcode.BoundText = ""
xDoc_no.Tag = DefineMode

xUserName2.Caption = ""
XTIME1.Caption = ""
XTIME2.Caption = ""
xDate.text = Format(Date, "DD-MM-YYYY")
xStore.BoundText = "1"
XNOTES.text = ""
xtime.Caption = Format(Time, "hh:nn")
grid1.Rows = 1
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
xDoc_no.Tag = 0

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

xDoc_no.Enabled = (nMode = DefineMode)
xDoc_no.Tag = nMode

End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_no
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xDoc_no.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload True
End Sub
Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
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
            If XORDER_NO.text <> "" Then
                If GetDesca("SELECT ITEM FROM FILE6_50 WHERE DOC_NO = " & MyParn(XORDER_NO.text) & " AND ITEM = " & (sitem), con) = "" Then
                    MsgBox " ’„› €Ì— „”Ã· »«·ÿ·»Ì…  "
                End If
            End If
            grid1.TextMatrix(Row, 2) = loctable!MOSM
            grid1.TextMatrix(Row, 3) = GetDesca("select desca from fact where code = " & MyParn(loctable!Fact), con)
            grid1.TextMatrix(Row, 5) = loctable!modelfact0 & ""
            grid1.TextMatrix(Row, 5) = loctable!DESCA
            grid1.TextMatrix(Row, 6) = loctable!scal
            grid1.TextMatrix(Row, 7) = loctable!Color
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
Dim GrdArray(5, 1)
Set Generalarray(0) = Me
'                                   0               1                   2                   3                                   4
Generalarray(1) = "SELECT   file6_55H.DOC_NO, file6_55H.ORDER_NO ,  FILE4_10.DESCA ,FILE6_55H.INV_NO  ,file6_55H.date, CONVERT(VARCHAR(10), file6_55H.date, 111) " & _
                  " FROM    file6_55H LEFT JOIN FILE4_10 ON FILE4_10.CODE = file6_55H.CODE "
Generalarray(2) = "Order by doc_no  DESC "
Generalarray(3) = 6000
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„ -  «—ÌŒ"
listarray(0, 1) = "( %%ORDER_NO%% OR %%INV_NO_NO%% OR ##date## )"
   

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1200

GrdArray(1, 0) = "—ﬁ„ ÿ·»Ì… "
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«·„Ê—œ "
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "—ﬁ„ «·›« Ê—… "
GrdArray(3, 1) = 1500

GrdArray(4, 0) = "«· «—ÌŒ"
GrdArray(4, 1) = 0

GrdArray(5, 0) = "«· «—ÌŒ"
GrdArray(5, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
'oSearchDoc.sId = "sales" & IIf(bopt3, "1", "0")
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub
Private Sub xDoc_No_Validate(Cancel As Boolean)
If xDoc_no.text = "" Then Cancel = True
End Sub
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

Private Sub xStore_GotFocus()
    myGotFocus xStore
End Sub
Private Sub xstore_LostFocus()
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
    con.Execute "Delete  From file6_55 where id = " & grid1.TextMatrix(nRow, grid1.Cols - 1)
End If
grid1.RemoveItem nRow
con.CommitTrans
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Function itemPrice(cItem) As Single
    itemPrice = GetDesca("select PRICE from file1_10 where item = " & MyParn(cItem), con)
End Function
Private Function myreplaceGrd(nRow) As Boolean
On Error GoTo myerror
Dim aInsert(3, 1), nTQ As Double
With grid1
    For i = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_no.text)
        
        aInsert(1, 0) = "item"
        aInsert(1, 1) = addstring(grid1.TextMatrix(i, 1))
        
        aInsert(2, 0) = "quant"
        aInsert(2, 1) = Val(.TextMatrix(i, 8))

        aInsert(3, 0) = "row"
        aInsert(3, 1) = i
        
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, "file6_55")
        Else
            con.Execute CreateUpdate(aInsert, "file6_55", " where ID = " & grid1.TextMatrix(i, .Cols - 1))
    End If
'''''    If XORDER_NO.Text = "" Then
'''''        If nRow > 0 Then
'''''            nTQ = Val(GetDesca("SELECT SUM(TREC-TORDER) FROM Q_ORDER_AMAZON WHERE doc_no = " & MyParn(xpo_no.Text) & " AND ITEM = " & Val(.TextMatrix(nRow, 1))) & "")
'''''            If nTQ > 0 Then
'''''                MsgBox " ≈” ·«„ ﬂ„Ì… «ﬂ»— „‰ «·„ÿ·Ê» "
'''''                .Select nRow, 1
'''''                .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbRed
'''''            End If
'''''        End If
'''''    End If
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
     cString = "SELECT file6_55.ITEM, FILE1_10.MOSM, FACT.DESCA , FILE1_10.MODELFACT0, FILE1_10.DESCA, FILE1_10.COLOR, FILE1_10.SCAL  , file6_55.Quant, BARCODE_GS1 , file6_55.ID " & _
          " FROM file6_55 INNER JOIN FILE1_10 ON file6_55.ITEM = FILE1_10.ITEM INNER JOIN FACT ON file1_10.FACT = FACT.CODE"
    cString = cString & " WHERE DOC_NO = " & MyParn(xDoc_no.text)
    cString = cString & " ORDER BY file6_55.ID"
    Set DATA11.Recordset = myRecordSet(cString, con)
    .AddItem ""
End With
Fixgrd
grid1.ShowCell grid1.Rows - 1, 1
End Sub
Private Sub xNotes_GotFocus()
myGotFocus XNOTES
End Sub
Private Sub xCode_GotFocus()
myGotFocus xcode
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_no
End Sub
Private Sub xdate_GotFocus()
myGotFocus xDate
End Sub
Private Function mysave(Optional bEnd As Boolean = True, Optional bPrint As Boolean = True) As Boolean
If Not MYVALID Then Exit Function
If Not myreplace Then Exit Function
If bEnd And lSave Then
    Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
    mydefine
Else
    openCardTable
    myUndo
End If
End Function
Private Sub myUndo()
'''On Error GoTo myerror
If CardTable.BOF And CardTable.EOF Then
    mydefine
Else
    If xDoc_no.text <> "" Then
        CardTable.Find "doc_no = " & MyParn(xDoc_no.text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cFilter = ""
cString = "SELECT * FROM file6_55H  "
If sDoc_no <> "" Then
    cFilter = cFilter & turn(cFilter, " AND ") & " DOC_NO = " & MyParn(sDoc_no)
End If
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY DOC_NO "
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = True) As Boolean
With grid1
If Not MYVALID(bIgMsg) Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 1)) Then Exit Function
End With
validRow = True
End Function
Private Sub HandleCntEdit()
xDoc_no.Tag = LoadMode
xDoc_no.Enabled = False
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
myLostFocus XNOTES
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
End Sub
Private Sub XORDER_NO_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And xcode.BoundText <> "" Then
    Dim Generalarray(5)
    Dim listarray(1, 4)
    Dim GrdArray(5, 1)
    
    Set Generalarray(0) = Me
'                               0       1               2       3               4
    Generalarray(1) = "SELECT  DOC_NO,[DATE] AS DATE1,DATE, FILE4_10.Desca , STORE_BR.desca , ISEND_ORDER" & _
                      " FROM  FILE6_50H INNER JOIN FILE4_10 ON FILE6_50H.CODE  = FILE4_10.CODE inner join STORE_BR on STORE_BR.CODE = FILE6_50H.STORE WHERE ISEND_ORDER = 0 "
    Generalarray(1) = Generalarray(1) & " AND FILE6_50H.CODE = " & MyParn(xcode.BoundText)
    Generalarray(2) = "Order by DATE DESC , doc_no desc "
    Generalarray(3) = 6000
    Generalarray(5) = True
    
    listarray(0, 0) = "«·—ﬁ„-≈”„ " & cCodeDesca & "-«· «—ÌŒ"
    listarray(0, 1) = "(inv_no Like '%cFilter%' or Doc_No Like '%cFilter%' or  FILE4_10.DESCA LIKE '%cFilter%' OR " & _
                      "##date##)"
    
    GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
    GrdArray(0, 1) = 1000
    
    GrdArray(1, 0) = "«· «—ÌŒ"
    GrdArray(1, 1) = 0
    
    GrdArray(2, 0) = "«· «—ÌŒ"
    GrdArray(2, 1) = 1500
    
    GrdArray(3, 0) = "≈”„ " & cCodeDesca
    GrdArray(3, 1) = 3000
    
    GrdArray(4, 0) = "«·„Œ“‰"
    GrdArray(4, 1) = 1500
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Load oSearchDoc
    oSearchDoc.Caption = "«” ⁄·«„"
    oSearchDoc.Show 1
End If
End Sub
Private Sub doprint()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
cStr1 = " SELECT FACT.DESCA AS fact_name, FILE1_10.MODELFACT0, FILE1_10.desca AS desca_item, FILE1_10.COLOR, FILE1_10.SCAL, FILE1_10.C_SCAL, FILE6_55.QUANT, FILE1_10.COST , FILE1_10.BARCODE_GS1 , FILE1_10.BARCODE , FILE1_10.BARCODE13 " & _
        " FROM   FILE1_10 INNER JOIN FILE6_55 ON FILE1_10.ITEM = FILE6_55.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE where file6_55.DOC_NO = " & MyParn(xDoc_no.text)
sourcetable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText


With sourcetable
Do While Not .EOF
    If !Quant > 0 Then
        temptable.AddNew
        temptable!str6 = TurnValue(xcode.text)
        temptable!str3 = TurnValue(xcode.text)
        temptable!Date1 = xDate.text
        temptable!str4 = !fact_name
        temptable!str10 = !modelfact0
        temptable!str2 = !desca_item
        If Not IsNull(!BARCODE_GS1) Then
            temptable!str9 = !BARCODE_GS1
        ElseIf Not IsNull(!BARCODE13) Then
            temptable!str9 = !BARCODE13
        ElseIf Not IsNull(!BARCODE) Then
            temptable!str9 = !BARCODE
        End If
        temptable!str8 = xDoc_no.text
        
        temptable!str2 = !desca_item
        temptable!STR7 = !Color
        temptable!str5 = !scal
        temptable!val5 = !C_SCAL
        temptable!VAL1 = !Quant
        temptable!str11 = TurnValue(xDoc_no.text)
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
main.Report1.ReportFileName = App.Path & "\Reports\Recevd_DOC.rpt"
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub

