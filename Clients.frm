VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Clients 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  ⁄„·«¡ √Ã·"
   ClientHeight    =   8490
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   9780
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
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   9780
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   450
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   7470
      Width           =   3570
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2700
         TabIndex        =   45
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
         Picture         =   "Clients.frx":0000
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "Clients.frx":21A7
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1800
         TabIndex        =   46
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
         Picture         =   "Clients.frx":41EE
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "Clients.frx":62D9
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   945
         TabIndex        =   47
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
         Picture         =   "Clients.frx":82D3
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "Clients.frx":A3E4
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   48
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
         Picture         =   "Clients.frx":C3DE
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "Clients.frx":E602
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   3285
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   45
      Width           =   6360
      Begin Threed.SSCommand cmdSave 
         Height          =   510
         Left            =   3195
         TabIndex        =   37
         Top             =   135
         Width           =   1005
         _ExtentX        =   1773
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
         Picture         =   "Clients.frx":106D3
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "Clients.frx":130C8
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   135
         Width           =   1005
         _ExtentX        =   1773
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
         Picture         =   "Clients.frx":15961
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   1095
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   135
         Width           =   1005
         _ExtentX        =   1773
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
         Picture         =   "Clients.frx":17C84
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "Clients.frx":1A420
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   510
         Left            =   2145
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   135
         Width           =   1005
         _ExtentX        =   1773
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
         Picture         =   "Clients.frx":1C8B4
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "Clients.frx":1EAF5
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   510
         Left            =   4245
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   135
         Width           =   1005
         _ExtentX        =   1773
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
         Picture         =   "Clients.frx":20DE2
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "Clients.frx":22DEA
      End
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   5295
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   135
         Width           =   1005
         _ExtentX        =   1773
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
         Picture         =   "Clients.frx":24DA1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "Clients.frx":2716C
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   405
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   1485
      Width           =   2400
      Begin VB.CheckBox xIsOnest 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "⁄„Ì· «„«‰« "
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
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   630
         Width           =   1275
      End
      Begin VB.CheckBox xIsMonth 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Ã„Ì⁄ ‘Â—Ï ··›Ê« Ì—"
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
         Height          =   360
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   225
         Width           =   1995
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1725
      Left            =   2835
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   810
      Width           =   6855
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
         Left            =   225
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   4695
      End
      Begin VB.TextBox xCode 
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
         Left            =   3600
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
      Begin VB.TextBox xManager 
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
         Left            =   225
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   4695
      End
      Begin MSDataListLib.DataCombo xAcc 
         Height          =   360
         Left            =   225
         TabIndex        =   3
         Top             =   1260
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Õ”«» "
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
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1305
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "≈”„ «·‘—ﬂ…"
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
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   945
         Width           =   840
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
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«”„"
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
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   615
         Width           =   375
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   1440
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   450
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   6030
      Width           =   9285
      Begin VB.TextBox xRemark 
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
         Height          =   465
         Left            =   45
         MaxLength       =   200
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   900
         Width           =   7305
      End
      Begin VB.TextBox xf_Balance 
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
         Left            =   225
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   180
         Width           =   2130
      End
      Begin VB.TextBox xf_Date 
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
         Left            =   5805
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox xDisc 
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
         Left            =   5805
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   540
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "„·«ÕŸ« "
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
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   945
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ √Ê · «·„œ…"
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
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   270
         Width           =   1200
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "—’Ìœ √Ê· «·„œ…"
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
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   225
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… «·Œ’„"
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
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   645
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   3525
      Left            =   450
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   2520
      Width           =   9285
      Begin VB.TextBox xid_no 
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
         Left            =   900
         MaxLength       =   14
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   3105
         Width           =   6450
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   375
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   180
         Width           =   330
      End
      Begin VB.TextBox xEMAIL 
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
         Left            =   900
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   2745
         Width           =   6450
      End
      Begin VB.TextBox xFAx 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   900
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2340
         Width           =   6450
      End
      Begin VB.TextBox xPhone1 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   900
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1935
         Width           =   6450
      End
      Begin VB.TextBox xAddress 
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
         Height          =   510
         Left            =   855
         MaxLength       =   100
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1395
         Width           =   6495
      End
      Begin MSDataListLib.DataCombo xgroup 
         Height          =   360
         Left            =   2970
         TabIndex        =   6
         Top             =   180
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xprice_list 
         Height          =   360
         Left            =   2970
         TabIndex        =   8
         Top             =   585
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xMan 
         Height          =   360
         Left            =   2970
         TabIndex        =   9
         Top             =   990
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„‰œÊ»"
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
         Left            =   7425
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   1035
         Width           =   600
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁ«∆„… «·”⁄—"
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
         Left            =   7425
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·»ÿ«ﬁ…"
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
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   3150
         Width           =   825
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «· ”ÃÌ· «·÷—Ì»Ï"
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
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   2745
         Width           =   1590
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "„Ê»Ì·"
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
         Left            =   7425
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   2385
         Width           =   420
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«· ·Ì›Ê‰"
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
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   2025
         Width           =   600
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ã„Ê⁄…"
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
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄‰Ê«‰"
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
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1395
         Width           =   540
      End
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   0
      Top             =   720
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
      Height          =   375
      Left            =   0
      TabIndex        =   49
      Top             =   8115
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   661
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   50
         Top             =   45
         Width           =   3195
         _ExtentX        =   5636
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
         Left            =   3240
         TabIndex        =   51
         Top             =   45
         Width           =   3855
         _ExtentX        =   6800
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
         Left            =   7110
         TabIndex        =   52
         Top             =   45
         Width           =   4365
         _ExtentX        =   7699
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
End
Attribute VB_Name = "Clients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public myFlag As Integer
Dim formMode As Byte, cTableName As String, cGroupname As String
Dim cFilter As String
Dim oSearch As New Search_abd
Dim CardTable As New ADODB.RecordSet
Const LoadMode = 1, DefineMode = 2
Private Sub Command3_Click()
Dim oFlagfrm As New flag_mainfrm, sBound As String
sBound = xgroup.BoundText
oFlagfrm.sTable = "FILE3_50"
oFlagfrm.sCaption = "„Ã„Ê⁄… «·„Ê—œÌ‰"
oFlagfrm.nZero = -1
oFlagfrm.bEdit = True
oFlagfrm.Show 1
DATA1.RecordSet.Requery
xgroup.BoundText = sBound
If Not xgroup.MatchedWithList Then xgroup.BoundText = ""
End Sub
Private Sub Form_Load()
bEdit = True

OpenCon con

Set DATA1.RecordSet = myRecordSet("SELECT * FROM file3_50", con)
Set xgroup.RowSource = DATA1
xgroup.ListField = "Desca"
xgroup.BoundColumn = "Code"

Set DATA2.RecordSet = myRecordSet("SELECT * FROM ACC0_10 ORDER BY CODE ", con)
Set xAcc.RowSource = DATA2
xAcc.ListField = "Desca"
xAcc.BoundColumn = "Code"

Set xprice_list.RowSource = myRecordSet("SELECT * FROM FILE6_53H ORDER BY DOC_NO", con)
xprice_list.ListField = "NOTES"
xprice_list.BoundColumn = "DOC_NO"

Set xMan.RowSource = myRecordSet("SELECT * FROM FILE6_25 ORDER BY CODE", con)
xMan.ListField = "DESCA"
xMan.BoundColumn = "CODE"

If Not openCardTable Then myDefine
End Sub
Private Sub CmdAdd_Click()
myDefine
xCode.text = RetZero(NewFlag("file3_10", "code", con), 4)
xCode.SetFocus
End Sub
Private Sub CmdDel_Click()
If MsgBox("Õ–› ?", vbOKCancel + vbDefaultButton2 + vbCritical) <> vbOK Then Exit Sub
On Error GoTo myError
con.Execute "delete from FILE3_10 where code = " & MyParn(xCode.text)
AddLod_Data cusername, 2, " Õ–› ⁄„·«¡ ", con, xCode.text, , , xDescA.text
If Not openCardTable(tbMode.tbPrevious, xCode.text) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
If Not bIgMsg Then Inform " „ «·Õ›Ÿ »‰Ã«Õ"
If sCode <> "" Then
    Unload Me
    Exit Sub
End If
If Not openCardTable(tbMode.tbFind, xCode.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub CmdUndo_Click()
myUndo
End Sub
Private Sub myUndo()
If xCode.Tag = DefineMode Then
    If Not openCardTable Then
        CmdAdd_Click
    End If
Else
    If Not openCardTable(tbMode.tbFind, xCode.text) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
End If
End Sub
Private Sub CmdInform_Click()
CLIENTLOOKUP Me, oSearch
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
Sub Handlecontrols(ByVal nMode As Integer)
xCode.Tag = nMode

Dim nRecord As Long, nRecords As Long
retRecords xCode.text, nRecords, nRecord

bEditRecord = bEdit

cmdSave.Enabled = bEditRecord
cmdAdd.Enabled = nMode = LoadMode And bEdit

cmddel.Enabled = bEditRecord And nMode = LoadMode

cmdNext.Enabled = nRecord < nRecords And nRecords <> 0 And nMode = LoadMode
cmdPrevious.Enabled = nRecord <> 1 And nRecords <> 0 And nMode = LoadMode
cmdLast.Enabled = nRecord < nRecords And nRecords <> 0 And nMode = LoadMode
cmdFirst.Enabled = nRecord <> 1 And nRecords <> 0 And nMode = LoadMode

If nMode = LoadMode Then
    panel1(0).Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
    panel1(0).Caption = "”Ã· ÃœÌœ (" & (nRecords + 1) & ")"
End If
End Sub
Private Function openCardTable(Optional pMode As Integer = tbMode.tblast, Optional pDoc_no As String = "", Optional bDefine As Boolean = True) As Boolean
Dim cString As String
Dim cWhere As String
Dim cOrder As String
Dim nRecords As Long, nRecord As Long

If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pDoc_no) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass

'On Error GoTo myerror
cFilter = ""


cString = "SELECT TOP 1 FILE3_10.* " & _
          " FROM FILE3_10"

If pMode = tbMode.tbFirst Then
    cOrder = "Order by FILE3_10.CODE"
ElseIf pMode = tbMode.tblast Then
    cOrder = "Order by FILE3_10.CODE DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "FILE3_10.CODE = " & MyParn(pDoc_no)
ElseIf pMode = tbMode.tbPrevious Then
    cWhere = "FILE3_10.CODE < " & MyParn(pDoc_no)
    cOrder = "Order by FILE3_10.CODE desc"
ElseIf pMode = tbMode.tbNext Then
    cWhere = "FILE3_10.CODE > " & MyParn(pDoc_no)
    cOrder = "Order by FILE3_10.CODE"
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
myError:
Me.MousePointer = vbNormal
MsgBox Err.Description
Err.Clear
End Function
Sub myDefine()
xCode.text = ""
xDescA.text = ""
xManager.text = ""
xEMAIL.text = ""
xgroup.BoundText = ""
xAcc.BoundText = ""
xIsOnest.Value = 0
xIsMonth.Value = 0
xAddress.text = ""
xprice_list.BoundText = ""
xPhone1.text = ""
xFAx.text = ""

xf_Date.text = ""
xf_Balance.text = ""
xDisc.text = ""
xid_no.text = ""
xMan.BoundText = ""
xRemark.text = ""
Handlecontrols DefineMode
End Sub
Sub myload()
xCode.text = CardTable!code & ""
xDescA.text = CardTable!DESCA
xManager.text = CardTable!Manager & ""
xMan.BoundText = CardTable!MAN & ""
xAddress.text = CardTable!Address & ""
xEMAIL.text = CardTable!email & ""
xPhone1.text = CardTable!PHONE1 & ""
xFAx.text = CardTable!Fax & ""
xIsOnest.Value = IIf(CardTable!ISONEST, 1, 0)
xIsMonth.Value = IIf(CardTable!IsMonth, 1, 0)
xf_Date.text = Format(CardTable!F_DATE, "dd-mm-yyyy")
xf_Balance.text = Format(CardTable!F_Balance, "##0.00")
xDisc.text = Format(CardTable!disc, "##0.00")
xgroup.BoundText = CardTable!Group & ""
xprice_list.BoundText = CardTable!price_list & ""
xAcc.BoundText = CardTable!ACC & ""
xRemark.text = CardTable!remark & ""
xid_no.text = CardTable!id_no & ""
xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "DESCA", addstring(xDescA.text))
aInsert = AddFlag(aInsert, "MANAGER", addstring(xManager.text))
aInsert = AddFlag(aInsert, "ADDRESS", addstring(xAddress.text))
aInsert = AddFlag(aInsert, "PHONE1", addstring(xPhone1.text))
aInsert = AddFlag(aInsert, "FAX", addstring(xFAx.text))
aInsert = AddFlag(aInsert, "f_balance", Val(xf_Balance.text))
aInsert = AddFlag(aInsert, "[GROUP]", addstring(xgroup.BoundText))
aInsert = AddFlag(aInsert, "[MAN]", addstring(xMan.BoundText))
aInsert = AddFlag(aInsert, "[price_list]", addvalue(xprice_list.BoundText))
aInsert = AddFlag(aInsert, "F_DATE", addDate(xf_Date.text))
aInsert = AddFlag(aInsert, "[email]", addstring(xEMAIL.text))
aInsert = AddFlag(aInsert, "[DISC]", Val(xDisc.text))
aInsert = AddFlag(aInsert, "[REMARK]", addstring(xRemark.text))
aInsert = AddFlag(aInsert, "[ISONEST]", xIsOnest.Value)
aInsert = AddFlag(aInsert, "[ISMONTH]", xIsMonth.Value)
aInsert = AddFlag(aInsert, "[ACC]", addstring(xAcc.BoundText))
aInsert = AddFlag(aInsert, "[ID_NO]", addstring(xid_no.text))

On Error GoTo myError
If xCode.Tag = DefineMode Then
    xCode.text = RetZero(NewFlag("FILE3_10", "code", con), 4)
    aInsert = AddFlag(aInsert, "CODE", addstring(xCode.text))
    con.Execute addInsert(aInsert, "FILE3_10")
Else
    con.Execute addUpdate(aInsert, "FILE3_10", "code = " & addstring(xCode.text))
End If
myreplace = True
Exit Function
myError:
MsgBox Err.Description
Err.Clear
End Function
Sub myProc()
openCardTable tbMode.tbFind, oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
Unload oSearch
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Err.Clear
End Sub
Private Sub xcode_LostFocus()
myLostFocus xCode
If Trim(xCode.text) = "" Then
     If xCode.Tag = LoadMode Then
        myDefine
    End If
Else
    xCode.text = RetZero(xCode.text)
'    If (Not (CardTable.EOF)) And xCode.Tag = LoadMode Then
'        If CardTable!CODE = xCode.text Then
'            Exit Sub
'        End If
'    End If
'
    If Not openCardTable(tbMode.tbFind, xCode.text) Then
       If xCode.Tag = LoadMode Then
            myDefine
       End If
    End If
End If
End Sub
Function MYVALID() As Boolean
If xCode.text = "" Then
    MsgBox "«·ﬂÊœ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If xgroup.BoundText = "" Then
    MsgBox "«·„Ã„Ê⁄… ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If xDescA.text = "" Then
    MsgBox "«·≈”„ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If Trim(xid_no.text) <> "" Then
    If Len(Trim(xid_no.text)) <> 14 Then
        MsgBox "—ﬁ„ «·»ÿ«ﬁ… €Ì— ’ÕÌÕ…"
        Exit Function
    End If
End If
MYVALID = True
End Function
Private Function retRecords(pDoc_no, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New ADODB.RecordSet
If pDoc_no <> "" Then
    cString = "SELECT Count(*) AS records,COUNT(CASE WHEN CODE <= " & MyParn(pDoc_no) & " THEN 1 END) AS record"
Else
    cString = "SELECT Count(*) AS records,0 as record"
End If

cString = cString & " FROM FILE3_10 " & turn(cFilter, " WHERE ") & cFilter
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
Private Sub xcode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xManager_GotFocus()
myGotFocus xManager
End Sub
Private Sub xManager_LostFocus()
myLostFocus xManager
End Sub
Private Sub xAcc_GotFocus()
myGotFocus xAcc
End Sub
Private Sub xAcc_LostFocus()
myLostFocus xAcc
If Not xAcc.MatchedWithList Then xAcc.BoundText = ""
End Sub
Private Sub XREMARK_GotFocus()
myGotFocus xRemark
End Sub
Private Sub XREMARK_LostFocus()
myLostFocus xRemark
End Sub
Private Sub xf_Balance_GotFocus()
myGotFocus xf_Balance
End Sub
Private Sub xf_Balance_LostFocus()
myLostFocus xf_Balance
End Sub
Private Sub xf_Date_GotFocus()
myGotFocus xf_Date
End Sub
Private Sub xf_Date_LostFocus()
myLostFocus xf_Date
myValidDate xf_Date
End Sub
Private Sub xDisc_GotFocus()
myGotFocus xDisc
End Sub
Private Sub xDisc_LostFocus()
myLostFocus xDisc
End Sub
Private Sub xid_no_GotFocus()
myGotFocus xid_no
End Sub
Private Sub xid_no_LostFocus()
myLostFocus xid_no
End Sub
Private Sub xEMAIL_GotFocus()
myGotFocus xEMAIL
End Sub
Private Sub xEMAIL_LostFocus()
myLostFocus xEMAIL
End Sub
Private Sub xFAx_GotFocus()
myGotFocus xFAx
End Sub
Private Sub xFAx_LostFocus()
myLostFocus xFAx
End Sub
Private Sub xPhone1_GotFocus()
myGotFocus xPhone1
End Sub
Private Sub xPhone1_LostFocus()
myLostFocus xPhone1
End Sub
Private Sub xAddress_GotFocus()
myGotFocus xAddress
End Sub
Private Sub xAddress_LostFocus()
myLostFocus xAddress
End Sub
Private Sub xGroup_GotFocus()
myGotFocus xgroup
End Sub
Private Sub xgroup_LostFocus()
myLostFocus xgroup
If Not xgroup.MatchedWithList Then xgroup.BoundText = ""
End Sub
Private Sub xMan_GotFocus()
myGotFocus xMan
End Sub
Private Sub xMan_LostFocus()
myLostFocus xMan
If Not xMan.MatchedWithList Then xMan.BoundText = ""
End Sub

