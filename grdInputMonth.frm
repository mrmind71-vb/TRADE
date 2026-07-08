VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form grdMonthInput 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈Ã„«·Ì «—»«Õ ‘Â—Ì… ··„Œ«“‰"
   ClientHeight    =   10290
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
   ScaleHeight     =   10290
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   1260
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   360
      Width           =   5685
      Begin VB.OptionButton Option1 
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
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "›—Ê⁄ ›ﬁÿ"
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
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   225
         Width           =   1140
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " ÊﬂÌ·«  ›ﬁÿ"
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   225
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   990
      Width           =   6900
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   979
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
         Picture         =   "grdInputMonth.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   1170
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   979
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
         Picture         =   "grdInputMonth.frx":2323
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "grdInputMonth.frx":4699
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   555
         Left            =   5760
         TabIndex        =   17
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "grdInputMonth.frx":681C
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   555
         Left            =   2295
         TabIndex        =   19
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "grdInputMonth.frx":96CF
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPdf 
         Height          =   555
         Left            =   3420
         TabIndex        =   18
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "grdInputMonth.frx":BBFA
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdSql 
         Height          =   555
         Left            =   4545
         TabIndex        =   45
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
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
         Picture         =   "grdInputMonth.frx":E0A7
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1680
      Left            =   6975
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   45
      Width           =   13335
      Begin VB.CheckBox XOutlet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Outlet "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4005
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   225
         Width           =   1095
      End
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
         Left            =   9900
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1815
      End
      Begin VB.TextBox xDate2 
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
         Left            =   8055
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1815
      End
      Begin VB.TextBox XMOSM2 
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
         Height          =   330
         Left            =   90
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1260
         Width           =   1635
      End
      Begin VB.TextBox xModelFact 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   5220
         MaxLength       =   15
         TabIndex        =   5
         Top             =   180
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   8055
         TabIndex        =   4
         Top             =   1260
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   330
         Left            =   8055
         TabIndex        =   3
         Top             =   900
         Width           =   3660
         _ExtentX        =   6456
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
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   330
         Left            =   8055
         TabIndex        =   2
         Top             =   555
         Width           =   3660
         _ExtentX        =   6456
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
      Begin MSDataListLib.DataCombo xage 
         Height          =   315
         Left            =   90
         TabIndex        =   10
         Top             =   180
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483643
         ListField       =   "&H80000005&"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xsex 
         Height          =   315
         Left            =   90
         TabIndex        =   11
         Top             =   540
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483643
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   330
         Left            =   3330
         TabIndex        =   7
         Top             =   540
         Width           =   3435
         _ExtentX        =   6059
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
      Begin MSDataListLib.DataCombo xStore 
         Height          =   330
         Left            =   3330
         TabIndex        =   8
         Top             =   900
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483643
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
      Begin MSDataListLib.DataCombo XSECTION 
         Height          =   330
         Left            =   3330
         TabIndex        =   9
         Top             =   1260
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin MSDataListLib.DataCombo xmosm20 
         Height          =   315
         Left            =   90
         TabIndex        =   12
         Top             =   900
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483643
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„Ê”„ œ«Œ·Ï"
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
         Index           =   0
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   900
         Width           =   1050
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·› —…"
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
         Left            =   11880
         TabIndex        =   36
         Top             =   225
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰Ê⁄ «·„Ê”„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1260
         Width           =   960
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "„Ê”„ «·„ÊœÌ·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   585
         Width           =   1230
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "«·›—⁄"
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
         Index           =   2
         Left            =   6885
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   945
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "«·ﬁ”„ "
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
         Index           =   1
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1350
         Width           =   690
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "—ﬁ„ „ÊœÌ·"
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
         Left            =   6855
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   180
         Width           =   765
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "›∆… ⁄„—Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   225
         Width           =   870
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·‰Ê⁄"
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
         Index           =   7
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   585
         Width           =   465
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„Ê—œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11820
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   540
         Width           =   645
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„Ã„Ê⁄… „ÊœÌ·« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   11820
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1305
         Width           =   1320
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„’‰⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11820
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   930
         Width           =   690
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   4770
      Top             =   -135
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   3285
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Height          =   7080
      Left            =   90
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1755
      Width           =   20220
      _cx             =   35666
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
      Cols            =   26
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   -405
      Top             =   1170
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   90
      Top             =   -135
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   330
      Left            =   0
      Top             =   585
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   -135
      Top             =   990
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   45
      Top             =   90
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   -45
      Top             =   405
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc data9 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      TabIndex        =   38
      Top             =   9825
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   39
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
         TabIndex        =   40
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
         TabIndex        =   41
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
         TabIndex        =   42
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
         TabIndex        =   43
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   44
      Top             =   9630
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "grdMonthInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim FieldTable As ADODB.Recordset
Dim cString As String, aHeader()
Private Sub CMD_EXEL_Click()
End Sub
Private Sub CMD_PRINT_Click()
End Sub
Private Sub cmdExcel_Click()
Me.MousePointer = 11
Fixgrd False
Dim aRow As Variant
aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aRow = AddFlag(aRow, aSub)
If grid1.Rows > 1 Then
    aSub = AddFlag(Empty, "row", 1)
    aSub = AddFlag(aSub, "bold", True)
    aSub = AddFlag(aSub, "word_wrap", False)
    aSub = AddFlag(aSub, "back_color", 40)
    aRow = AddFlag(aRow, aSub)
End If
ToFileExelNew grid1, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption, retHeader(aHeader, 0, 4), retHeader(aHeader, 4, 4), retHeader(aHeader, 8, 4))
Fixgrd True
Me.MousePointer = 0
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
If Not MYVALID Then Exit Sub
Me.MousePointer = vbHourglass
myload
Me.MousePointer = vbNormal
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xDate1.text) Then
    MsgBox " «—ÌŒ «·„œ… «·«Ê· €Ì— „Õœœ"
    Exit Function
End If
If Not IsDate(xDate2.text) Then
    MsgBox " «—ÌŒ «·„œ… «·À«‰Ì €Ì— „Õœœ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub cmdPdf_Click()
myPrint "1", True
End Sub
Private Sub cmdPrint_Click()
myPrint
End Sub

Private Sub cmdSql_Click()
myload True
End Sub

Private Sub Form_Load()
openCon con

Set DATA6.Recordset = myCmd("Select Code,DescA From File0_81  order by Desca", con)
Set xsex.RowSource = DATA6
xsex.ListField = "Desca"
xsex.BoundColumn = "Code"

Set DATA7.Recordset = myCmd("Select Code,DescA From STORE_ALL WHERE " & cUserStore, con)
Set xStore.RowSource = DATA7
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"
If cBranch <> "00" Then
    xStore.BoundText = cBranchStore
    xStore.Enabled = False
End If

Set data1.Recordset = myCmd("Select Code,DescA From File1_10SC ORDER BY code ", con)
Set XSECTION.RowSource = data1
XSECTION.ListField = "Desca"
XSECTION.BoundColumn = "Code"

Set data1.Recordset = myCmd("SELECT CODE,DESCA FROM FILE0_40 UNION ALL SELECT CODE,DESCA FROM BRANCH WHERE STORE IS NULL  " & cWhere, con)
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

Set DATA2.Recordset = myCmd("Select Code,DescA From File4_10 order by Desca", con)
Set xSupp.RowSource = DATA2
xSupp.ListField = "Desca"
xSupp.BoundColumn = "Code"

Set DATA3.Recordset = myCmd("Select Code,DescA From File1_50 ORDER BY DESCA", con)
Set xGroup.RowSource = DATA3
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

Set DATA4.Recordset = myCmd("Select mosm ,descA From mosm ORDER BY date DESC ", con)
Set xMosm.RowSource = DATA4
xMosm.ListField = "Desca"
xMosm.BoundColumn = "MOSM"
xMosm.BoundText = cPMosm


Set DATA5.Recordset = myCmd("Select code ,desca From fact ORDER BY DESCA ", con)
Set xFact.RowSource = DATA5
xFact.ListField = "Desca"
xFact.BoundColumn = "Code"

Set DATA6.Recordset = myCmd("Select Code,DescA From File0_81  order by Desca", con)
Set xsex.RowSource = DATA6
xsex.ListField = "Desca"
xsex.BoundColumn = "Code"

Set DATA7.Recordset = myCmd("Select Code,DescA From STORE_ALL WHERE " & cUserStore, con)
Set xStore.RowSource = DATA7
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"
If cBranch <> "00" Then
    xStore.BoundText = cBranchStore
    xStore.Enabled = False
End If

Set data8.Recordset = myCmd("Select Code,Desca From File0_82  order by Desca", con)
Set xage.RowSource = data8
xage.ListField = "Desca"
xage.BoundColumn = "Code"

Set data9.Recordset = myCmd("Select 'S' AS Code UNION ALL Select 'W' UNION ALL Select 'M'", con)
Set xmosm20.RowSource = data9
xmosm20.ListField = "code"
xmosm20.BoundColumn = "Code"

Set grid1.DataSource = DATA11

Fixgrd
End Sub
Private Sub myload(Optional bString As Boolean)
With grid1
ReDim aHeader(13)

Dim aPrm As Variant
Dim aPrm2 As Variant
If IsDate(xDate1.text) Then
    aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xDate1.text))
    aPrm2 = AddFlag(aPrm2, "DATE1", addDate(xDate1.text))
    aHeader(0) = BetweenString(xDate1.text, xDate2.text)
End If

If IsDate(xDate2.text) Then
    aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xDate2.text))
    aPrm2 = AddFlag(aPrm2, "DATE2", addDate(xDate2.text))
    aHeader(1) = BetweenString(xDate1.text, xDate2.text)
End If

If XSECTION.MatchedWithList Then
    aPrm = AddFlag(aPrm, "SECTION", XSECTION.BoundText)
    aPrm2 = AddFlag(aPrm2, "SECTION", XSECTION.BoundText)
    aHeader(2) = "«·ﬁ”„ : " & XSECTION.text
Else
    aPrm = AddFlag(aPrm, "SECTION", XSECTION.BoundText)
    aPrm2 = AddFlag(aPrm2, "SECTION", addvalue(XSECTION.BoundText))
    aHeader(2) = "«·ﬁ”„ : " & XSECTION.text
End If

If xSupp.MatchedWithList Then
    aPrm = AddFlag(aPrm, "SUP", xSupp.BoundText)
    aPrm2 = AddFlag(aPrm2, "SUP", addstring(xSupp.BoundText))
    aHeader(3) = "«·„Ê—œ : " & xSupp.text
End If

If xGroup.MatchedWithList Then
    aPrm = AddFlag(aPrm, "GROUP", xGroup.BoundText)
    aPrm2 = AddFlag(aPrm2, "GROUP", addstring(xGroup.BoundText))
    aHeader(4) = "„Ã„Ê⁄… «·’‰› : " & xGroup.text
End If

If xFact.MatchedWithList Then
    aPrm = AddFlag(aPrm, "FACT", xFact.BoundText)
    aPrm2 = AddFlag(aPrm2, "FACT", addstring(xFact.BoundText))
    aHeader(5) = "«·„’‰⁄ : " & xFact.text
End If

If XOutlet.Value = "1" Then
    aPrm = AddFlag(aPrm, "outlet", XOutlet.Value)
    aPrm2 = AddFlag(aPrm2, "outlet", XOutlet.Value)
    aHeader(6) = "Outlet"
End If

If xMosm.MatchedWithList Then
    aPrm = AddFlag(aPrm, "MOSM", xMosm.BoundText)
    aPrm2 = AddFlag(aPrm2, "MOSM", addstring(xMosm.BoundText))
    aHeader(7) = "«·„Ê”„ : " & xMosm.text
End If

If xmosm20.MatchedWithList Then
    aPrm = AddFlag(aPrm, "MOSM20", xmosm20.BoundText)
    aPrm2 = AddFlag(aPrm2, "MOSM20", addstring(xmosm20.BoundText))
    aHeader(8) = "„Ê”„ œ«Œ·Ì : " & xmosm20.text
End If

If xage.MatchedWithList Then
    aPrm = AddFlag(aPrm, "ModelAge", xage.BoundText)
    aPrm2 = AddFlag(aPrm2, "ModelAge", addstring(xage.BoundText))
    aHeader(9) = "«·›∆…  «·⁄„—Ì… : " & xage.text
End If

If xsex.MatchedWithList Then
    aPrm = AddFlag(aPrm, "ModelSex", xsex.BoundText)
    aPrm2 = AddFlag(aPrm2, "ModelAge", addstring(xsex.BoundText))
    aHeader(10) = "«·‰Ê⁄ : " & xsex.text
End If

If xStore.MatchedWithList Then
    aPrm = AddFlag(aPrm, "STORE", xStore.BoundText)
    aPrm2 = AddFlag(aPrm2, "STORE", addstring(xStore.BoundText))
    aHeader(11) = "„Œ“‰ : " & xStore.text
End If

If Trim(XMOSM2.text) <> "" Then
    If XMOSM2.text = "W" Then
        aPrm = AddFlag(aPrm, "MOSM2", "'W','M'")
        aPrm2 = AddFlag(aPrm2, "MOSM2", "'W','M'")
    Else
        aPrm = AddFlag(aPrm, "MOSM2", XMOSM2.text)
        aPrm2 = AddFlag(aPrm2, "MOSM2", addstring(XMOSM2.text))
    End If
    aHeader(12) = "‰Ê⁄ «·„Ê”„ : " & XMOSM2.text
End If

If Option1(1).Value Or Option1(1).Value Then
    aPrm = AddFlag(aPrm, "ISBRANCH", IIf(Option1(1).Value, "1", "0"))
    aPrm2 = AddFlag(aPrm2, "ISBRANCH", IIf(Option1(1).Value, "1", "0"))
    aHeader(13) = IIf(Option1(1).Value, Option1(1).Caption, Option1(2).Caption)
End If


If Not bOpt5 Then
    aPrm = AddFlag(aPrm, "usercode", TurnValue(nusercode))
    aPrm2 = AddFlag(aPrm2, "usercode", TurnValue(nusercode))
End If

Dim cString As String
cString = myPrcString("dbo.sp_items_month", aPrm2)

If bString Then
    Clipboard.Clear
    Clipboard.SetText cString
    Exit Sub
End If

'Set DATA11.Recordset = myCmd("dbo.sp_items_month", con, adStoredProc, aPrm, 600)
Set DATA11.Recordset = myCmd(cString, con, , , 600)
End With
Fixgrd
panel1(0).Caption = "⁄œœ «·”Ã·«  : " & DATA11.Recordset.RecordCount
End Sub
Sub Fixgrd(Optional bGroup As Boolean = True)
With grid1
.RowHeight(0) = 1000
.WordWrap = True
.TextMatrix(0, 0) = "«· «—ÌŒ"
.TextMatrix(0, 1) = "«·„Ê—œ"
.TextMatrix(0, 2) = "«·„’‰⁄"
.TextMatrix(0, 3) = "„Ê”„"
.TextMatrix(0, 4) = "„ÊœÌ·"
.TextMatrix(0, 5) = "«·’‰›"
.TextMatrix(0, 6) = "«·›—⁄"

.TextMatrix(0, 7) = "ﬂ„Ì… Ê«—œ"
.TextMatrix(0, 8) = "‰ﬂ·›… «·Ê«—œ"
.TextMatrix(0, 9) = " ﬁÌ„ «·Ê«—œ „” Â·ﬂ"

.TextMatrix(0, 10) = "ﬂ„Ì… ’«œ—"
.TextMatrix(0, 11) = " ﬂ·›… ’«œ—"
.TextMatrix(0, 12) = " ﬁÌ„ ’«œ— „” Â·ﬂ"

.TextMatrix(0, 13) = "’«›Ï «·ﬂ„Ì…"
.TextMatrix(0, 14) = "’«›Ï ‰ﬂ·›… «·ﬁÌ„…"
.TextMatrix(0, 15) = "’«›Ï „” Â·ﬂ "

.TextMatrix(0, 16) = "ﬂ„Ì… „»Ì⁄« "
.TextMatrix(0, 17) = " ﬂ·›… „»Ì⁄« "
.TextMatrix(0, 18) = "ﬁÌ„… „»Ì⁄« "
.TextMatrix(0, 19) = "—»Õ „»Ì⁄« "

.TextMatrix(0, 20) = "«·‰Ê⁄"
.TextMatrix(0, 21) = "«·›∆… «·⁄„—Ì…"
.TextMatrix(0, 22) = "«·ﬁ”„"
.TextMatrix(0, 23) = "„Ã„Ê⁄… «·„ÊœÌ·« "
.TextMatrix(0, 24) = "„Ã„Ê⁄… „ÊœÌ·"
.TextMatrix(0, 25) = "„Ê”„ œ«Œ·Ï"

.ColWidth(1) = 1200
.ColWidth(2) = 1200
.ColWidth(3) = 700
.ColWidth(4) = 1500
.ColWidth(5) = 3000
.ColWidth(6) = 2000

.ColWidth(7) = 1100
.ColWidth(8) = 1600
.ColWidth(9) = 1600

.ColWidth(10) = 1100
.ColWidth(11) = 1600
.ColWidth(12) = 1600

.ColWidth(13) = 1100
.ColWidth(14) = 1600
.ColWidth(15) = 1600

.ColWidth(16) = 1100
.ColWidth(17) = 1600
.ColWidth(18) = 1600
.ColWidth(19) = 1600
    
.ColWidth(20) = 1600
.ColWidth(21) = 1600
.ColWidth(22) = 1600
.ColWidth(23) = 1600

For Col = 7 To 15 Step 3
    .ColFormat(Col) = "##,##"
    .ColFormat(Col + 1) = "##,##.##"
    .ColFormat(Col + 2) = "##,##.##"
    .ColDataType(Col) = flexDTDouble
Next

.ColFormat(16) = "##,##"
.ColFormat(17) = "##,##.##"
.ColFormat(18) = "##,##.##"
.ColFormat(19) = "##,##.##"

.ExplorerBar = flexExSort
.Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4


.SubtotalPosition = flexSTAbove
If bGroup Then
    For Col = 7 To 19
        .Subtotal flexSTSum, 0, Col, "##,##", &HC0FFC0, vbBlack, True, "≈Ã„«·Ì %s"
    Next
End If

.SubtotalPosition = flexSTAbove
For Col = 7 To 19
    .Subtotal flexSTSum, -1, Col, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
Next
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set grdChargefrm1 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
myload
End Sub
Private Sub myPrint(Optional pDevice As String = "", Optional bIgPreview As Boolean = False)
If grid1.Rows < 3 Then Exit Sub
Dim aRow As Variant, aSub As Variant

For i = 1 To grid1.Rows - 2
    If grid1.TextMatrix(i, 2) = "" Then
        aSub = AddFlag(Empty, "row", i)
        aSub = AddFlag(aSub, "col", 1)
        aSub = AddFlag(aSub, "cols", 2)
        aSub = AddFlag(aSub, "text", grid1.TextMatrix(i, 1))
        aRow = AddFlag(aRow, aSub)
    End If
Next

aSub = AddFlag(Empty, "row", grid1.Rows - 1)
aSub = AddFlag(aSub, "col", 1)
aSub = AddFlag(aSub, "cols", 2)
aSub = AddFlag(aSub, "text", "≈Ã„«·Ì")
aRow = AddFlag(aRow, aSub)

'PrintGrdNew.sDeviceType = pDevice
'PrintGrdNew.Vp.Preview = Not bIgPreview

'PrintGrdNew.bIgPreview = bIgPreview
PrintGrdNew.doprint Me.grid1, 0.95, 0, "ÃÊ‰ÌÊ—", Me.Caption, retHeader(aHeader, 0, 2), , False, True, 10, , aRow, Array(1)
If Not bIgPreview Then

    PrintGrdNew.Show 1
Else
    Unload PrintGrdNew
    'PrintGrdNew.Show 1
    'Unload PrintGrdNew
End If
End Sub


Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub XMOSM2_GotFocus()
myGotFocus XMOSM2
End Sub
Private Sub XMOSM2_LostFocus()
myLostFocus XMOSM2
End Sub
Private Sub xModelFACT_GotFocus()
myGotFocus xModelFact
End Sub
Private Sub XModelFACT_LostFocus()
myLostFocus xModelFact
End Sub
Private Sub xGroup_GotFocus()
myGotFocus xGroup
End Sub
Private Sub xgroup_LostFocus()
myLostFocus xGroup
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub
Private Sub xFact_GotFocus()
myGotFocus xFact
End Sub
Private Sub xFACT_LostFocus()
myLostFocus xFact
If Not xFact.MatchedWithList Then xFact.BoundText = ""
End Sub
Private Sub xSupp_GotFocus()
myGotFocus xSupp
End Sub
Private Sub xSupp_LostFocus()
myLostFocus xSupp
If Not xSupp.MatchedWithList Then xSupp.BoundText = ""
End Sub
Private Sub xage_GotFocus()
myGotFocus xage
End Sub
Private Sub xage_LostFocus()
myLostFocus xage
If Not xage.MatchedWithList Then xage.BoundText = ""
End Sub
Private Sub xsex_GotFocus()
myGotFocus xsex
End Sub
Private Sub xsex_LostFocus()
myLostFocus xsex
If Not xsex.MatchedWithList Then xsex.BoundText = ""
End Sub
Private Sub xMosm_GotFocus()
myGotFocus xMosm
End Sub
Private Sub xMosm_LostFocus()
myLostFocus xMosm
If Not xMosm.MatchedWithList Then xMosm.BoundText = ""
End Sub
Private Sub XSTORE_GotFocus()
myGotFocus xStore
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus xStore
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub xSection_GotFocus()
myGotFocus XSECTION
End Sub
Private Sub xSection_LostFocus()
myLostFocus XSECTION
If Not XSECTION.MatchedWithList Then XSECTION.BoundText = ""
End Sub
Private Sub xmosm20_GotFocus()
myGotFocus xmosm20
End Sub
Private Sub xmosm20_LostFocus()
myLostFocus xmosm20
If Not xmosm20.MatchedWithList Then xmosm20.BoundText = ""
End Sub
