VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form PrintDiscBarcode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· ﬂ—Ê  Œ’„ "
   ClientHeight    =   10935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15120
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "⁄—÷ ﬂ—Ê  „” Œœ„… ›ﬁÿ"
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
      Left            =   4635
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   9315
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "«ŸÂ«— «·ÃÂ…"
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
      Left            =   12915
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   4005
      Width           =   2130
   End
   Begin Threed.SSCommand CMD_MAKEBARCODE 
      Height          =   600
      Left            =   8910
      TabIndex        =   35
      Top             =   3465
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   1058
      _Version        =   196610
      CaptionStyle    =   1
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PrintDiscBarcode.frx":0000
      Caption         =   "«÷«›… «·»«—ﬂÊœ ·ﬂ—Ê  «·Œ’„"
      Alignment       =   1
      ButtonStyle     =   1
      PictureAlignment=   4
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   8910
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   45
      Width           =   6135
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   4725
         Picture         =   "PrintDiscBarcode.frx":240F
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   420
         Left            =   3180
         MaskColor       =   &H00FFFFFF&
         Picture         =   "PrintDiscBarcode.frx":4BE2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   420
         Left            =   1635
         MaskColor       =   &H00FFFFFF&
         Picture         =   "PrintDiscBarcode.frx":718E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "PrintDiscBarcode.frx":9A28
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   7335
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   3375
      Width           =   1500
      Begin VB.CommandButton CmdUndo 
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "PrintDiscBarcode.frx":9B72
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   1410
      End
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "PrintDiscBarcode.frx":C0EB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   870
      ScaleWidth      =   15120
      TabIndex        =   13
      Top             =   9765
      Width           =   15120
      Begin VB.Frame Frame9 
         Height          =   870
         Left            =   8955
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   0
         Width           =   4155
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
            TabIndex        =   43
            Top             =   -270
            Visible         =   0   'False
            Width           =   105
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
            TabIndex        =   42
            Top             =   135
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
            TabIndex        =   41
            Top             =   135
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
            TabIndex        =   40
            Top             =   495
            Width           =   1815
         End
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
      End
      Begin VB.Frame Frame4 
         Height          =   555
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   -45
         Width           =   3975
         Begin VB.CommandButton cmdNext 
            Height          =   375
            Left            =   2985
            Picture         =   "PrintDiscBarcode.frx":C52D
            Style           =   1  'Graphical
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "«· «·Ì"
            Top             =   135
            Width           =   915
         End
         Begin VB.CommandButton cmdPrevious 
            Height          =   375
            Left            =   2070
            Picture         =   "PrintDiscBarcode.frx":EAED
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "«·”«»ﬁ"
            Top             =   135
            Width           =   915
         End
         Begin VB.CommandButton cmdLast 
            Height          =   375
            Left            =   915
            Picture         =   "PrintDiscBarcode.frx":110C0
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "√ŒÌ—"
            Top             =   135
            Width           =   915
         End
         Begin VB.CommandButton cmdFirst 
            Height          =   375
            Left            =   45
            Picture         =   "PrintDiscBarcode.frx":1379A
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "√Ê·"
            Top             =   135
            Width           =   870
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   12
      Top             =   10635
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "01:18 „"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   2490
      Left            =   7380
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   810
      Width           =   7575
      Begin VB.TextBox xDoc_No 
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
         Height          =   360
         Left            =   4230
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1425
      End
      Begin VB.TextBox XVALUE 
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
         Height          =   360
         Left            =   3960
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1449
         Width           =   1695
      End
      Begin VB.TextBox XDATE2 
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
         Height          =   360
         Left            =   3960
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1890
         Width           =   1695
      End
      Begin VB.TextBox xsales 
         Alignment       =   2  'Center
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
         Height          =   360
         Left            =   135
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   630
         Width           =   330
      End
      Begin VB.TextBox XCOUNT 
         Alignment       =   2  'Center
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
         Height          =   360
         Left            =   135
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1845
         Width           =   1335
      End
      Begin VB.TextBox XDESCA 
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
         Height          =   360
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   990
         Width           =   5520
      End
      Begin VB.TextBox xDate 
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
         Height          =   360
         Left            =   135
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1425
      End
      Begin MSDataListLib.DataCombo XPALCE 
         Height          =   315
         Left            =   2175
         TabIndex        =   2
         Top             =   615
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Õœ «œ‰Ï „»Ì⁄«  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1530
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "⁄œœ «·ﬂ—Ê  : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1935
         Width           =   1020
      End
      Begin VB.Label Label10 
         Caption         =   "«·ÃÂ… :  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5805
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   675
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "ﬁÌ„… «·Œ’„ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5895
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1530
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "»Ì«‰ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5895
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1125
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «‰ Â«¡ «·Œ’„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5940
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1935
         Width           =   1410
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5850
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   225
         Width           =   930
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   3600
      Top             =   7650
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Height          =   9150
      Left            =   0
      TabIndex        =   14
      Top             =   135
      Width           =   7260
      _cx             =   12806
      _cy             =   16140
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
   Begin Threed.SSCommand CMD_PRINT 
      Height          =   555
      Left            =   12825
      TabIndex        =   27
      Top             =   3420
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   979
      _Version        =   196610
      CaptionStyle    =   1
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PrintDiscBarcode.frx":15E79
      Caption         =   "ÿ»«⁄… »«—ﬂÊœ ··ﬂ—Ê "
      Alignment       =   1
      ButtonStyle     =   1
      PictureAlignment=   4
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   3510
      Top             =   7470
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   4830
      Left            =   7335
      TabIndex        =   44
      Top             =   4455
      Width           =   7710
      _cx             =   13600
      _cy             =   8520
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
Attribute VB_Name = "PrintDiscBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim cFile As String, cFileHeader As String, sName As String
Dim formMode
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim aInsert(9, 1)

aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xdoc_no.text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = addDate(xDate.text)

aInsert(2, 0) = "DESCA"
aInsert(2, 1) = addstring(XDESCA.text)

aInsert(3, 0) = "PLACE"
aInsert(3, 1) = addstring(XPALCE.BoundText)

aInsert(4, 0) = "COUNT"
aInsert(4, 1) = Val(XCOUNT.text)

aInsert(5, 0) = "DATE2"
aInsert(5, 1) = addDate(xdate2.text)

aInsert(6, 0) = "VALUE"
aInsert(6, 1) = Val(XVALUE.text)

aInsert(7, 0) = "sales"
aInsert(7, 1) = Val(xsales.text)

On Error GoTo myerror
con.BeginTrans
If xdoc_no.Enabled Then
    
    aInsert(8, 0) = "USERNAME1"
    aInsert(8, 1) = addstring(xUserName1.Caption)
    
    aInsert(9, 0) = "TIME1"
    aInsert(9, 1) = "GETDATE()"
    
    con.Execute CreateInsert(aInsert, "CARD_DISCH")
Else
    aInsert(8, 0) = "USERNAME2"
    aInsert(8, 1) = addstring(xUserName2.Caption)
    
    aInsert(9, 0) = "TIME2"
    aInsert(9, 1) = "GETDATE()"
    
    con.Execute CreateUpdate(aInsert, "CARD_DISCH", " where doc_no = " & addstring(xdoc_no.text))

End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
    CardTable.Find "doc_No = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    myload
    Unload Search3
End Sub
Private Sub Check2_Click()
    Fixgrd1
End Sub

Private Sub CMD_MAKEBARCODE_Click()
Dim cBarCode As String
If MsgBox("  ≈÷«›… »«—ﬂÊœ ﬂ—Ê  «·Œ’„  ", vbYesNo + vbDefaultButton2) Then
    If Val(GetDesca("SELECT COUNT(*) FROM CARD_DISC WHERE DOC_NO = " & MyParn(xdoc_no.text), con) & "") > 0 Then
'        MsgBox " ÌÊÃœ »«—ﬂÊœ „”Ã· „‰ ﬁ»·  "
        MsgBox " ”Ì „ «” ﬂ„«· «·»«—ﬂÊœ «·€Ì— „”Ã· „‰ ﬁ»·  "
'        Exit Sub
    End If
    For i = 1 To Val(XCOUNT.text)
        Me.Caption = " BARCODE " & cBarCode
        cBarCode = xdoc_no.text & RetZero(i, 4)
        nRow = grid1.FindRow(cBarCode, 1, 0)
        If nRow = -1 Then
            con.Execute " INSERT INTO CARD_DISC  ( DOC_NO , BARCODE_DISC ) VALUES (" & addstring(xdoc_no.text) & " ," & addstring(cBarCode) & ")"
        End If
    Next i
    myloadgrd
    myloadgrd2
    MsgBox "  „ ⁄„· «·»«—ﬂÊœ "
End If
End Sub


Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
'   con.Execute "Delete  From CARD_DISC     where Doc_No = " & MyParn(xDoc_No.text)
    con.Execute "Delete  From CARD_DISCH    where Doc_No = " & MyParn(xdoc_no.text)
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
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdFirst_Click()
    CardTable.MoveFirst
    myload
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
cString = "SELECT Doc_No , Convert(Varchar(10),Date,111) , DESCA , VALUE " & _
          " FROM CARD_DISCH "
          
Generalarray(1) = cString
Generalarray(2) = " order by date "
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "„” ‰œ- «—ÌŒ "
listarray(0, 1) = "(%%DOC_NO%% or %%DESCA%% or " & _
                  " ##Date##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«·»Ì«‰"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "«·Œ’„"
GrdArray(3, 1) = 1200

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = ""
Search3.Show 1
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
myDefine
xdoc_no.SetFocus
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
myreplace
CardTable.Requery
CardTable.Find "Doc_No = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
myload
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.EOF And CardTable.BOF Then
    myDefine
Else
    If xdoc_no.Enabled Then CardTable.MoveLast Else CardTable.Find "Doc_No = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
    myload
End If
End Sub

Private Sub Command3_Click()
ReDim aLocal(6)
aLocal(0) = "DISC_PLACE"
aLocal(1) = "Code"
aLocal(2) = "Desca"
aLocal(3) = "ﬂÊœ "
aLocal(4) = "»Ì«‰ "
aLocal(5) = "ÃÂ«  ﬂ—Ê  «·Œ’„"
aLocal(6) = 2
FlagFrm2.bedit = True
FlagFrm2.myPublic = aLocal
FlagFrm2.Show 1
DATA2.Refresh

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
    openCon con
    
    
    data3.ConnectionString = strCon
    data3.RecordSource = "SELECT * FROM DISC_PLACE"
    Set XPALCE.RowSource = data3
    XPALCE.ListField = "Desca"
    XPALCE.BoundColumn = "Code"


    Set grid1.DataSource = data1
    Set GRID2.DataSource = DATA2

    Me.Caption = DocTitle

    Set CardTable = New ADODB.Recordset
    CardTable.Open "SELECT * FROM CARD_DISCH ORDER BY DOC_NO", con, adOpenStatic, adLockReadOnly, adCmdText

If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    myDefine
    Fixgrd1
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CardTable.Close
    Set CardTable = Nothing
    closeCon con
End Sub
Private Function MYVALID() As Boolean
If Trim(xdoc_no.text) = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If XPALCE.BoundText = "" Then
    MsgBox "«·ÃÂ… €Ì— ”·Ì„"
    Exit Function
End If

If Val(XCOUNT.text) = 0 Then
    MsgBox "⁄œœ €Ì— ”·Ì„ "
    Exit Function
End If

If Val(XVALUE.text) = 0 Then
    MsgBox "«·Œ’„ €Ì— ”·Ì„ "
    Exit Function
End If



MYVALID = True
End Function
Private Sub myload()
Dim GrdTable As New ADODB.Recordset
xdoc_no.text = CardTable!doc_no
xDate.text = Format(CardTable!Date, "dd-mm-yyyy")
xdate2.text = Format(CardTable!date2, "dd-mm-yyyy")
XPALCE.BoundText = TurnValue(CardTable!PLACE, Null, "")
XDESCA.text = CardTable!DESCA & ""
XVALUE.text = CardTable!Value & ""
xsales.text = CardTable!sales & ""
XCOUNT.text = CardTable!Count & ""

xUserName1.Caption = CardTable!UserName1 & ""
XTIME1.Caption = CardTable!TIME1 & ""

grid1.Rows = 1
GRID2.Rows = 1
myloadgrd
myloadgrd2

Handlecontrols LoadMode

End Sub
Private Sub myDefine()

xdoc_no.text = RetZero(Newflag("CARD_DISCH", "DOC_NO", con), 6)
xDate.text = Format(Date, "dd-mm-yyyy")
XPALCE.BoundText = ""
xUserName1.Caption = cusername
XTIME1.Caption = ""
XDESCA.text = ""
XVALUE.text = ""
xsales.text = ""
XCOUNT.text = ""
grid1.Rows = 1
GRID2.Rows = 1
'grid1.TextMatrix(grid1.Rows - 1, 0) = defBox

Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
    cmdNewInv.Enabled = (nMode = LoadMode)
    cmdFirst.Enabled = (nMode = LoadMode)
    cmdLast.Enabled = (nMode = LoadMode)
    cmdNext.Enabled = (nMode = LoadMode)
    CmdDelInv.Enabled = (nMode = LoadMode)
    cmdPrevious.Enabled = (nMode = LoadMode)
    xdoc_no.Enabled = (nMode = DefineMode)
'    CmdSave.Enabled = (grid1.Rows = 1)
    
End Sub

Private Sub VSFlexGrid1_Click()

End Sub

Private Sub xDoc_No_LostFocus()
If Trim(xdoc_no.text) = "" Then Exit Sub
xdoc_no.text = RetZero(xdoc_no.text)
CardTable.Find "Doc_no = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub xDoc_No_Validate(Cancel As Boolean)
    If xdoc_no.text = "" Then Cancel = True
End Sub
Sub myloadgrd()
    Dim cString As String
'                       0                               1                       2                   3               4                                                                   5                   6
    cString = " SELECT CARD_DISC.BARCODE_DISC  ,T_SALESDOC_ALL.DOC_NO2      , T_SALESDOC_ALL.DATE , STORE_BR.DESCA , T_SALESDOC_ALL.T_total , T_SALESDOC_ALL.discount , T_SALESDOC_ALL.T_total-T_SALESDOC_ALL.discount FROM CARD_DISC LEFT JOIN T_SALESDOC_ALL ON T_SALESDOC_ALL.DOC_NO = CARD_DISC.SALES_DOC_NO LEFT JOIN STORE_BR ON STORE_BR.CODE = T_SALESDOC_ALL.STORE  WHERE CARD_DISC.DOC_NO = " & MyParn(xdoc_no.text) & " ORDER BY BARCODE_DISC "
    Set data1.Recordset = myRecordSet(cString, con)
    Fixgrd1
End Sub
Sub Fixgrd1()
With grid1
    .Cols = 7
    .ColWidth(0) = 1400
    .ColWidth(1) = 800
    .ColWidth(2) = 1200
    .ColWidth(3) = 1200
    .ColWidth(4) = 700
    .ColWidth(5) = 700
    .ColWidth(6) = 700
    
    .TextMatrix(0, 0) = "»«—ﬂÊœ"
    .TextMatrix(0, 1) = "„” ‰œ »Ì⁄"
    .TextMatrix(0, 2) = "«· «—Ì€"
    .TextMatrix(0, 3) = "«·›—⁄"
    .TextMatrix(0, 4) = "«·ﬁÌ„…"
    .TextMatrix(0, 5) = "«·ŒÀ„"
    .TextMatrix(0, 6) = "’«›Ï"
    For nRow = 1 To .Rows - 1
        If Check2.Value <> 0 Then
            .RowHidden(nRow) = (.TextMatrix(nRow, 1) = "")
        Else
            .RowHidden(nRow) = False
        End If
    Next nRow
End With
End Sub
Private Sub CMD_PRINT_Click()
    aPrinterBar = AddFlag(Empty, "name", RetPrinterByType("2"))
    Doprint_BarCodeDisc
    Set myForm = Me
    CardPrintNew_2.Show 1

End Sub

Private Function Doprint_BarCodeDisc() As Boolean
nLastMargin = 0
SettingArray(cUpMargin) = MyMeasure(0.2)
SettingArray(cLeftMargin) = MyMeasure(0)
SettingArray(cCardWidth) = MyMeasure(4)
SettingArray(cCardHeight) = MyMeasure(2)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1

SettingArray(cRows) = 1
SettingArray(cCols) = 1

'SettingArray(cPageWidth) = MyMeasure(21)

Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM card"
temptable.Open "card", contemp, adOpenStatic, adLockOptimistic, adCmdTable


cString = " Select CARD_DISCH.DOC_NO , CARD_DISCH.COUNT ,CARD_DISCH.VALUE , CARD_DISCH.DATE2 , CARD_DISCH.SALES , CARD_DISC.BARCODE_DISC , CARD_DISCH.DESCA  " & _
          " FROM CARD_DISCH INNER JOIN CARD_DISC ON CARD_DISCH.DOC_NO = CARD_DISC.DOC_NO LEFT JOIN DISC_PLACE ON DISC_PLACE.CODE = CARD_DISCH.PLACE   WHERE CARD_DISCH.DOC_NO = " & MyParn(xdoc_no.text)
cString = cString & "  order by CARD_DISC.BARCODE_DISC"

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ

Dim sourcetable As New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If

For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = NCOLS, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    blastrow = (nRow = NROWS)
    temptable.AddNew
    temptable!CardNo = nCard
    temptable.Update
Next
'«‰ Â«¡

nCard = 0
nRow = 0
nCard = 0
nCol = 0
NCOLS = SettingArray(cCols)
NROWS = SettingArray(cRows)
nFixLeft = MyMeasure(0.1)
nFixlast = MyMeasure(0.18)
nFixlastbar = MyMeasure(0.1)

With sourcetable
Do Until .EOF
'************
    For i = 1 To 1
        nCard = nCard + 1
        nCol = IIf(nCol = NCOLS, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        
        nHeight = 0
        For nCount = 1 To 1
            blastrow = (nRow = NROWS And nCount = 2)
            
            temptable.AddNew
            temptable!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            temptable!Top = MyMeasure(0.1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Width = MyMeasure(3.5)
            temptable!Height = MyMeasure(0.4)
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 8
            temptable!TextAlign = taCenterTop
            temptable!text = sourcetable!BARCODE_DISC
            temptable!isBarcode = True
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            temptable!Top = MyMeasure(0.5) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Width = MyMeasure(3.5)
            temptable!Height = MyMeasure(0.4)
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 8
            temptable!TextAlign = taCenterTop
            temptable!text = sourcetable!BARCODE_DISC
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
            
            
            temptable.AddNew
            temptable!Left = MyMeasure(0.1) - IIf(blastcol, nLastCol, 0)
            temptable!Top = MyMeasure(0.75) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Width = MyMeasure(3.7)
            temptable!Height = MyMeasure(0.9)
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontSize = 7
            temptable!TextAlign = taCenterTop
            temptable!text = "⁄ œ ‘—«∆ﬂ »„‘ —Ì«  »ﬁÌ„… " & Format(sourcetable!sales, "#0") & "  Õ’· ⁄·Ï  Œ’„ »ﬁÌ„… " & Format(sourcetable!Value, "#0") & " Õ Ï " & Format(sourcetable!date2, "DD-MM-YYYY")
            temptable!isBarcode = False
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
            
            
            temptable.AddNew
            temptable!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            temptable!Top = MyMeasure(1.5) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Width = MyMeasure(3.7)
            temptable!Height = MyMeasure(0)
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 7
            temptable!TextAlign = taCenterTop
            temptable!text = "«·ﬁ”„Ì… ’«·Õ… ··„‰ Ã«  Œ«—Ã «·Œ’„ ›ﬁÿ "
            temptable!isBarcode = False
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            temptable!Top = MyMeasure(1.9) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Width = MyMeasure(3.5)
            temptable!Height = MyMeasure(0)
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 8
            temptable!TextAlign = taCenterTop
            If Check1.Value <> 0 Then temptable!text = sourcetable!DESCA
            temptable!isBarcode = False
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
            
            
            
            
            nHeight = SettingArray(cCardHeight) / 2
        Next
' ----------------
    Next
    sourcetable.MoveNext
Loop
End With
Set sourcetable = Nothing
End Function
Sub myloadgrd2()
    Dim cString As String
'                       0                               1                       2                   3               4                                                                   5                   6
    cString = " SELECT STORE_BR.DESCA , COUNT(*) , SUM(T_SALESDOC_ALL.T_total ) , SUM(T_SALESDOC_ALL.discount ), SUM(T_SALESDOC_ALL.T_total-T_SALESDOC_ALL.discount) , SUM(T_SALESDOC_ALL.T_total-T_SALESDOC_ALL.discount-T_SALESDOC_ALL.T_COST)   FROM CARD_DISC LEFT JOIN T_SALESDOC_ALL ON T_SALESDOC_ALL.DOC_NO = CARD_DISC.SALES_DOC_NO INNER JOIN STORE_BR ON STORE_BR.CODE = T_SALESDOC_ALL.STORE  WHERE CARD_DISC.DOC_NO = " & MyParn(xdoc_no.text) & " GROUP BY STORE_BR.CODE , STORE_BR.DESCA ORDER BY  STORE_BR.CODE "
    Set DATA2.Recordset = myRecordSet(cString, con)
    Fixgrd2
End Sub

Sub Fixgrd2()
With GRID2
    .Cols = 6
    .ColWidth(0) = 1600
    .ColWidth(1) = 800
    .ColWidth(2) = 1200
    .ColWidth(3) = 1200
    .ColWidth(4) = 1200
    .ColWidth(5) = 1200
    
    .TextMatrix(0, 0) = "«·›—⁄"
    .TextMatrix(0, 1) = "⁄œœ »Ê‰« "
    .TextMatrix(0, 2) = "«·ﬁÌ„…"
    .TextMatrix(0, 3) = "«·Œ’„"
    .TextMatrix(0, 4) = "’«›Ï"
    .TextMatrix(0, 5) = "«·—»Õ"
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 1, "#0", vbYellow, , True, " "
    .Subtotal flexSTSum, -1, 2, "#0", vbYellow, , True, " "
    .Subtotal flexSTSum, -1, 3, "#0", vbYellow, , True, " "
    .Subtotal flexSTSum, -1, 4, "#0", vbYellow, , True, " "
    .Subtotal flexSTSum, -1, 5, "#0", vbYellow, , True, " "
End With
End Sub

