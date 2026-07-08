VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form BoxModelBarCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· »Ê‰«  Œ’„ "
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10125
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
   ScaleHeight     =   7725
   ScaleWidth      =   10125
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CheckBox xNoPrint 
      Alignment       =   1  'Right Justify
      Caption         =   "»œÊ‰ ÿ»«⁄… »«—ﬂÊœ ··’‰œÊﬁ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6030
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   3645
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   4365
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   45
      Width           =   5640
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   4095
         Picture         =   "BoxModelBarCode.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   420
         Left            =   2790
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BoxModelBarCode.frx":27D3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   420
         Left            =   1440
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BoxModelBarCode.frx":4D7F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BoxModelBarCode.frx":7619
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   4410
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3645
      Width           =   1500
      Begin VB.CommandButton CmdUndo 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BoxModelBarCode.frx":7763
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BoxModelBarCode.frx":9CDC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   510
      ScaleWidth      =   10125
      TabIndex        =   11
      Top             =   6915
      Width           =   10125
      Begin VB.Frame Frame4 
         Height          =   555
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   -45
         Width           =   3975
         Begin VB.CommandButton cmdNext 
            Height          =   375
            Left            =   2985
            Picture         =   "BoxModelBarCode.frx":A11E
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "«· «·Ì"
            Top             =   135
            Width           =   915
         End
         Begin VB.CommandButton cmdPrevious 
            Height          =   375
            Left            =   2070
            Picture         =   "BoxModelBarCode.frx":C6DE
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "«·”«»ﬁ"
            Top             =   135
            Width           =   915
         End
         Begin VB.CommandButton cmdLast 
            Height          =   375
            Left            =   915
            Picture         =   "BoxModelBarCode.frx":ECB1
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "√ŒÌ—"
            Top             =   135
            Width           =   915
         End
         Begin VB.CommandButton cmdFirst 
            Height          =   375
            Left            =   45
            Picture         =   "BoxModelBarCode.frx":1138B
            Style           =   1  'Graphical
            TabIndex        =   21
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
      TabIndex        =   10
      Top             =   7425
      Width           =   10125
      _ExtentX        =   17859
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
            TextSave        =   "02:21 „"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame6 
      Height          =   585
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   6165
      Visible         =   0   'False
      Width           =   3720
      Begin VB.TextBox xusername 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   135
         Width           =   3465
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2805
      Left            =   4410
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   765
      Width           =   5595
      Begin VB.TextBox XBOX 
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
         Left            =   3915
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   2295
         Width           =   525
      End
      Begin VB.TextBox XCOLOR 
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
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1476
         Width           =   1290
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
         Locked          =   -1  'True
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1044
         Width           =   3900
      End
      Begin VB.TextBox XQUANT 
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
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   2295
         Width           =   795
      End
      Begin VB.TextBox XCOUNT 
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
         Left            =   2070
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   2295
         Width           =   750
      End
      Begin VB.TextBox XITEM 
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
         Left            =   2610
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1875
      End
      Begin VB.TextBox XSCAL 
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
         Left            =   1507
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1476
         Width           =   750
      End
      Begin VB.TextBox XMODEL 
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
         Left            =   2340
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1476
         Width           =   1695
      End
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
         Left            =   2610
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1425
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
         TabIndex        =   2
         Top             =   180
         Width           =   1425
      End
      Begin MSDataListLib.DataCombo xtypebox 
         Height          =   315
         Left            =   135
         TabIndex        =   32
         Top             =   1908
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XSEX 
         Height          =   315
         Left            =   135
         TabIndex        =   41
         Top             =   630
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label10 
         Caption         =   "«·‰Ê⁄ : "
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   675
         Width           =   660
      End
      Begin VB.Label Label9 
         Caption         =   "⁄œœ ’‰«œÌﬁ"
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
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   2340
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·„ÊœÌ· :"
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
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1530
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "≈”„ «·„ÊœÌ· :"
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
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1125
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "‰Ê⁄ «· ⁄»∆…"
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
         Left            =   4230
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1935
         Width           =   810
      End
      Begin VB.Label Label3 
         Caption         =   "œ” … / ⁄·»…"
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
         Left            =   990
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   2385
         Width           =   1110
      End
      Begin VB.Label Label2 
         Caption         =   "⁄œœ «·⁄·» "
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
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   2385
         Width           =   930
      End
      Begin VB.Label Label4 
         Caption         =   "»«—ﬂÊœ :"
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
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   675
         Width           =   525
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
         TabIndex        =   8
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
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   930
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   5130
      Top             =   5265
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
      Height          =   6675
      Left            =   90
      TabIndex        =   12
      Top             =   90
      Width           =   4200
      _cx             =   7408
      _cy             =   11774
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
   Begin Threed.SSCommand CMD_PRINT 
      Height          =   690
      Left            =   7650
      TabIndex        =   25
      Top             =   3780
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   1217
      _Version        =   196610
      CaptionStyle    =   1
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
      Picture         =   "BoxModelBarCode.frx":13A6A
      Caption         =   "ÿ»«⁄…  ﬂ  «·⁄·» Ê «·’‰œÊﬁ"
      Alignment       =   1
      PictureAlignment=   3
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   5220
      Top             =   5670
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
   Begin MSAdodcLib.Adodc data3 
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
Attribute VB_Name = "BoxModelBarCode"
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

aInsert(2, 0) = "userName"
aInsert(2, 1) = addstring(sUserName)

aInsert(3, 0) = "item"
aInsert(3, 1) = addstring(xitem.text)

aInsert(4, 0) = "typebox"
aInsert(4, 1) = addstring(xtypebox.BoundText)

aInsert(5, 0) = "COUNT"
aInsert(5, 1) = Val(XCOUNT.text)

aInsert(6, 0) = "QUANT"
aInsert(6, 1) = Val(xquant.text)

aInsert(7, 0) = "BOX"
aInsert(7, 1) = Val(xBox.text)

aInsert(8, 0) = "NOPRINT"
aInsert(8, 1) = Val(xNoPrint.Value)

aInsert(9, 0) = "SEX"
aInsert(9, 1) = addstring(xsex.BoundText)

On Error GoTo myerror
con.BeginTrans
If xdoc_no.Enabled Then
    xdoc_no.text = RetZero(Newflag("BOX_H", "doc_no"))
    aInsert(0, 1) = addstring(xdoc_no.text)
    con.Execute CreateInsert(aInsert, "BOX_H")
    AddSubGrid
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub AddSubGrid()

For nCount = 1 To (Val(XCOUNT.text) * Val(xBox.text))
    Dim aInsert(2, 1)
    aInsert(0, 0) = "doc_no"
    aInsert(0, 1) = addstring(xdoc_no.text)
                
    aInsert(1, 0) = "row"
    aInsert(1, 1) = nCount
    
    con.Execute CreateInsert(aInsert, "BOX")
Next nCount
End Sub
Sub myProc()
    If ActiveControl.Name = xitem.Name Then
        xitem.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
        xitem_LostFocus
        Unload Search3
    Else
        CardTable.Find "doc_No = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
        myload
        Unload Search3
    End If
    
End Sub

Private Sub Cmd_Print_Click()
    On Error GoTo myerror
    If xdoc_no.Enabled Then Exit Sub
    DoPrintCover
    con.Execute " UPDATE box_H SET PRINTED = 1 , [TIME] = GETDATE() WHERE DOC_NO  = " & MyParn(xdoc_no.text)
    CardTable.Requery
    CardTable.Find "Doc_No = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
    myload
    Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From BOX   where Doc_No = " & MyParn(xdoc_no.text)
    con.Execute "Delete  From BOX_H where Doc_No = " & MyParn(xdoc_no.text)
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
Dim ListArray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
cString = "SELECT Doc_No , Convert(Varchar(10),Date,111) , BOX_H.item , file1_10.desca " & _
          " FROM BOX_H inner join FILE1_10 ON file1_10.item = box_h.item "
          
Generalarray(1) = cString
Generalarray(2) = " order by date "
Generalarray(3) = 4000
Generalarray(5) = False

ListArray(0, 0) = "„” ‰œ- «—ÌŒ "
ListArray(0, 1) = "(%%DOC_NO%% or %%BOX_H.item%% or " & _
                  " ##Date##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "—ﬁ„ «·’‰›"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "«·’‰›"
GrdArray(3, 1) = 4000

searchArray = Array(Generalarray, ListArray, GrdArray)
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
If MsgBox("”Ê› Ì „ Õ›Ÿ «·„” ‰œ Ê «‰‘«¡  ﬂ  »«—ﬂÊœ ··⁄·» Ê «·’‰œÊﬁ", vbYesNo + vbDefaultButton2) = vbYes Then
    If Not myreplace Then Exit Sub
    Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
End If
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
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
    openCon con
    FixData1
    sName = "«·⁄„Ì·"
    cFile = "box"
    cFileHeader = "BOX_H"
    DocTitle = "ÿ»«⁄… »«—ﬂÊœ ··⁄·» Ê «·’‰«œÌﬁ"
    con.Execute " UPDATE BOX_H SET BOX = 1 WHERE BOX IS NULL "
    con.Execute " UPDATE BOX_H SET DOC_NO = '00000' + BOX_H.DOC_NO WHERE LEN(BOX_H.DOC_NO) = 1 "
    con.Execute " UPDATE BOX   SET DOC_NO = '00000' + BOX.DOC_NO WHERE LEN(BOX.DOC_NO) = 1 "
    
    con.Execute " UPDATE BOX_H SET DOC_NO = '0000' + BOX_H.DOC_NO WHERE LEN(BOX_H.DOC_NO) = 2 "
    con.Execute " UPDATE BOX   SET DOC_NO = '0000' + BOX.DOC_NO WHERE LEN(BOX.DOC_NO) = 2 "
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "SELECT * FROM typebox"
    Set xtypebox.RowSource = DATA2
    xtypebox.ListField = "Desca"
    xtypebox.BoundColumn = "Code"

    data3.ConnectionString = strCon
    data3.RecordSource = "SELECT * FROM SUB_SEX "
    Set xsex.RowSource = data3
    xsex.ListField = "Desca"
    xsex.BoundColumn = "Code"

    Set grid1.DataSource = data1
    data1.ConnectionString = strCon

    Me.Caption = DocTitle

    Set CardTable = New ADODB.Recordset
    CardTable.Open "SELECT * FROM BOX_H ORDER BY DOC_NO", con, adOpenStatic, adLockReadOnly, adCmdText


If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    myDefine
    Fixgrd
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

If xtypebox.BoundText = "" Then
    MsgBox "‰Ê⁄ «· ⁄»∆… €Ì— ”·Ì„"
    Exit Function
End If

If xitem.text = "" Then
    MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã· ’ÕÌÕ"
    Exit Function
End If

If XMODEL.text = "" Then
    MsgBox "«·„ÊœÌ· €Ì— „”Ã·"
    Exit Function
End If



MYVALID = True
End Function
Private Sub myload()
Dim GrdTable As New ADODB.Recordset
xdoc_no.text = CardTable!doc_no
xDate.text = Format(CardTable!Date, "dd-mm-yyyy")
xtypebox.BoundText = TurnValue(CardTable!typebox, Null, "")
xitem.text = CardTable!Item & ""
xquant.text = CardTable!Quant & ""
XCOUNT.text = CardTable!Count & ""
xBox.text = CardTable!BOX & ""
xsex.BoundText = CardTable!SEX & ""
xNoPrint.Value = IIf(CardTable!NOPRINT, 1, 0)
aret = aGetDesca("SELECT DESCA , MODEL , SCAL , COLOR FROM FILE1_10 WHERE ITEM = " & MyParn(xitem.text))
If UBound(aret) > 0 Then
    xDesca.text = aret(1)
    XMODEL.text = aret(2)
    xScal.text = aret(3)
    xColor.text = aret(4)
End If
xusername.text = TurnValue(CardTable!UserName, Null, "")
handleControls LoadMode
myloadgrd
If CardTable!Printed Then
    cmdSave.Enabled = False
    CmdDelInv.Enabled = False Or bOpt4
End If
End Sub
Private Sub myloadgrd()
With grid1
    cString = "SELECT ROW , SER_NO " & _
               " FROM BOX WHERE DOC_NO = " & MyParn(xdoc_no.text) & " Order by Row "
    data1.RecordSource = cString
    data1.Refresh
End With
Fixgrd
End Sub
Private Sub myDefine()
xdoc_no.text = RetZero(Val(Newflag("BOX", "doc_no")), 6)
xDate.text = Format(Date, "dd-mm-yyyy")
xtypebox.BoundText = ""
xusername.text = ""
xNoPrint.Value = 0
xDesca.text = ""
xitem.text = ""
xtypebox.text = ""
xsex.text = ""

xBox.text = "1"
xColor.text = ""
XCOUNT.text = ""
xquant.text = ""
xScal.text = ""
XMODEL.text = ""
grid1.Rows = 1
grid1.TextMatrix(grid1.Rows - 1, 0) = defBox
handleControls DefineMode
End Sub
Private Sub handleControls(nMode)
cmdNewInv.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
CmdDelInv.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xdoc_no.Enabled = (nMode = DefineMode)
cmdSave.Enabled = (grid1.Rows = 1)
End Sub

Private Sub grid1_EnterCell()
    grid1.Editable = flexEDNone
End Sub

Private Sub SSCommand1_Click()
    DoPrintCover
End Sub

Private Sub xdoc_no_LostFocus()
If Trim(xdoc_no.text) = "" Then Exit Sub
xdoc_no.text = RetZero(xdoc_no.text)
CardTable.Find "Doc_no = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub xDoc_No_Validate(Cancel As Boolean)
    If xdoc_no.text = "" Then Cancel = True
End Sub
Private Sub Fixgrd()
With grid1
    .Editable = flexEDKbdMouse
    .Cols = 2
    .FormatString = "„.|" & "—ﬁ„ «·⁄·»…"
    .ColWidth(0) = 1500
    .ColWidth(1) = 1500
End With
End Sub
Private Sub xItem_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then ItemsLookup
End Sub

Private Sub xitem_LostFocus()
Dim aret
If xitem.text <> "" Then
    aret = aGetDesca("SELECT DESCA , MODEL , SCAL , COLOR , SEX FROM FILE1_10 WHERE ITEM = " & MyParn(xitem.text))
    If UBound(aret) > 0 Then
        xDesca.text = aret(1)
        XMODEL.text = aret(2)
        xScal.text = aret(3)
        xColor.text = aret(4)
        xsex.BoundText = aret(5) & ""
    Else
        xDesca.text = ""
        XMODEL.text = ""
        xScal.text = ""
        xColor.text = ""
        xsex.BoundText = "    "
    End If
End If
End Sub

Private Sub xtypebox_LostFocus()
If xtypebox.BoundText <> "" Then
    xquant.text = Val(GetDesca("SELECT QUANT FROM TYPEBOX WHERE CODE = " & MyParn(xtypebox.BoundText)) & "")
    XCOUNT.text = Val(GetDesca("SELECT COUNT FROM TYPEBOX WHERE CODE = " & MyParn(xtypebox.BoundText)) & "")
End If
End Sub
Private Sub DoPrintCover()
Dim temptable As ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim aret
Dim nRowPrint As Double
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable


cStr1 = " SELECT * FROM BOX WHERE DOC_NO = " & MyParn(xdoc_no.text) & " ORDER BY ser_no "
sourcetable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText

If sourcetable.RecordCount = 0 Then
    MsgBox "·« ÌÊÃœ „ÊœÌ·« "
    Exit Sub
End If
aret = aGetDesca("SELECT DESCA , MODEL , SCAL , COLOR , model2 FROM FILE1_10 WHERE ITEM = " & MyParn(xitem.text))
With temptable
sourcetable.MoveFirst
Do While Not sourcetable.EOF
    nRowPrint = nRowPrint + 1
    .AddNew
    !str1 = "‘—ﬂ… ÃÊ‰ÌÊ— ··„·«»”"
    !str7 = TurnValue(xsex.text)
    !str2 = RetZero(sourcetable!Ser_no, 6)
    If UBound(aret) > 0 Then
        If aret(5) & "" = "" Then
            !str3 = XMODEL.text
        Else
            !str3 = aret(5)
        End If
        !str4 = xScal.text
         
        If Val(xquant.text) = 4 Then
            !str5 = ArbString(Val(xquant.text) & " ﬁÿ⁄…")
        Else
            !str5 = ArbString(Val(xquant.text) & " œ” …")
        End If
'        Select Case Val(XQUANT.Text)
'            Case 1
'                !str5 = "œ” …"
'            Case 2
'                !str5 = ArbString("2 œ” …")
'            Case 0.5
'                !str5 = "‰’ œ” …"
'        End Select
        !STR16 = "17130155103005"
        !STR6 = xDesca.text
        !str2 = RetZero(sourcetable!Ser_no)
        End If
    .Update
    
    If nRowPrint = Val(XCOUNT.text) Then nRowPrint = 0
    If nRowPrint = 0 And xNoPrint.Value = 0 Then
        .AddNew
        !str1 = "‘—ﬂ… ÃÊ‰ÌÊ— ··„·«»”"
        !str2 = RetZero(xdoc_no.text)
        If UBound(aret) > 0 Then
            If aret(5) & "" = "" Then
                !str3 = XMODEL.text
            Else
                !str3 = aret(5)
            End If
            !str4 = xScal.text
            !str5 = Val(xquant.text) & " œ” …"
            !str5 = ArbString(Val(XCOUNT.text) & "  X  " & Val(xquant.text) & " œ” …")
            
'            Select Case Val(XQUANT.Text)
'                Case 1
'                    !str5 = ArbString(Val(XCOUNT.Text) & "  X  " & "œ” …")
'                Case 2
'                    !str5 = ArbString(Val(XCOUNT.Text) & "  X  " & "2 œ” …")
'                Case 0.5
'                    !str5 = ArbString(Val(XCOUNT.Text) & "  X  " & "‰’ œ” …")
'            End Select
            !STR6 = xDesca.text
            End If
        .Update
    End If
    sourcetable.MoveNext
Loop
End With

contemp.BeginTrans
contemp.CommitTrans

Report1.Reset
FixPrinter Report1
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    Report1.WindowShowPrintSetupBtn = True
    Report1.ReportFileName = App.Path & "\Reports\Rep_Cover.rpt"
    Report1.DataFiles(0) = tempFile
    Report1.WindowState = crptMaximized
    Report1.Destination = crptToWindow
    Report1.Action = 1
End If
'temptable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Sub ItemsLookup()
Dim Generalarray(5)
Dim ListArray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select File1_10.item,File1_10.Desca,file1_10.COLOR , FILE1_10.SCAL From file1_10 WHERE FILE1_10.MODEL IS NOT NULL "
Generalarray(2) = "Order by file1_10.ITEM "
Generalarray(3) = 4200
Generalarray(5) = False

ListArray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
ListArray(0, 1) = "(FILE1_10.ITEM LIKE 'cFilter%' or  %%FILE1_10.DESCA%%) "


GrdArray(0, 0) = "ﬂÊœ «·’‰›"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "≈”„ «·’‰›"
GrdArray(1, 1) = 5000

GrdArray(2, 0) = "„ﬁ«”"
GrdArray(2, 1) = 1000

GrdArray(3, 0) = "«··Ê‰"
GrdArray(3, 1) = 1500


searchArray = Array(Generalarray, ListArray, GrdArray)
Search3.Caption = "«” ⁄·«„ «·«’‰«›"
Search3.Show 1
End Sub
Private Sub FixData1()
On Error Resume Next
Dim FS1 As ADODB.Command
Set FS1 = New ADODB.Command

cString = "ALTER TABLE [dbo].[BOX_H] ADD   [NOPRINT] [bit] NOT NULL CONSTRAINT [DF_BOX_H_NOPRINT]  DEFAULT ((0))"

FS1.CommandType = adCmdText
Set FS1.ActiveConnection = con
FS1.CommandText = cString
FS1.Execute
End Sub

