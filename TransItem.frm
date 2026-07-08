VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
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
   Begin VB.CommandButton Command1 
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
      Height          =   600
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   1215
      Width           =   1635
   End
   Begin VB.CheckBox xClosed 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   330
      Left            =   495
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   225
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Frame Frame8 
      Height          =   1050
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   765
      Width           =   1950
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
         Left            =   90
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   585
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
         Left            =   90
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   180
         Width           =   780
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
         Left            =   990
         RightToLeft     =   -1  'True
         TabIndex        =   40
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
         Left            =   990
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   630
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
            Key             =   ""
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
         Picture         =   "TransItem.frx":0000
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
         Picture         =   "TransItem.frx":2579
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
         Picture         =   "TransItem.frx":48DC
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
         Picture         =   "TransItem.frx":7176
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
         Picture         =   "TransItem.frx":9722
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
         Picture         =   "TransItem.frx":BEF5
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
      Left            =   1530
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Width           =   7260
      Begin VB.CommandButton cmdAddItems 
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
         Height          =   600
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   135
         Visible         =   0   'False
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
         Picture         =   "TransItem.frx":E361
         Caption         =   " ”ÃÌ· „ÊœÌ·«  "
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin Threed.SSCommand CMD_PRINT2 
         Height          =   600
         Left            =   1890
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
         Picture         =   "TransItem.frx":10B69
         Caption         =   "ÿ»«⁄… „” ‰œ «·Ã—œ"
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin Threed.SSCommand CMD_PRINT 
         Height          =   600
         Left            =   45
         TabIndex        =   35
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
         Picture         =   "TransItem.frx":12F78
         Caption         =   "ÿ»«⁄… «· ”ÊÌ…"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   7695
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   765
      Width           =   7485
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
      Height          =   6780
      Left            =   90
      TabIndex        =   4
      Top             =   1890
      Width           =   15060
      _cx             =   26564
      _cy             =   11959
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
         Picture         =   "TransItem.frx":15387
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
         Picture         =   "TransItem.frx":175D3
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
         Picture         =   "TransItem.frx":198B0
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
         Picture         =   "TransItem.frx":1BADD
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
         Picture         =   "TransItem.frx":1DD74
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
         Picture         =   "TransItem.frx":2012E
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
         Picture         =   "TransItem.frx":22530
         Caption         =   " «·€«¡  —ÕÌ·"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin VB.Frame Frame7 
      Height          =   690
      Left            =   2115
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   8685
      Width           =   5685
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   510
         Left            =   45
         TabIndex        =   32
         Top             =   135
         Visible         =   0   'False
         Width           =   5595
         _ExtentX        =   9869
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
   Begin VB.Label xbranch 
      Alignment       =   1  'Right Justify
      Caption         =   "Label7"
      Height          =   240
      Left            =   1755
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   1530
      Width           =   915
   End
End
Attribute VB_Name = "StockFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim Search31 As New Search3, search32 As New Search3, oSearchItem As New Search3
Dim ItemTable As New ADODB.Recordset

Dim CardTable As ADODB.Recordset
Dim cFilter As String
Dim cDefStore As String
Dim formMode, dDateLast As String
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional nRow As Long = -1) As Boolean
Dim aInsert(5, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = DateSq(xDate.Text)

aInsert(2, 0) = "store"
aInsert(2, 1) = addstring(xstore.BoundText)

aInsert(3, 0) = "DESCA"
aInsert(3, 1) = addstring(xdesca.Text)

aInsert(4, 0) = "ISDATE"
aInsert(4, 1) = "GETDATE()"

aInsert(5, 0) = "BRANCH"
aInsert(5, 1) = addstring(xbranch.Caption)

con.BeginTrans
On Error GoTo myerror
If xDoc_No.Tag = DefineMode Then
    xDoc_No.Text = NewflagBranch("FILE0_10H", "DOC_NO", cBranch)
    aInsert(0, 1) = addstring(xDoc_No.Text)
    con.Execute CreateInsert(aInsert, "FILE0_10H")
Else
    con.Execute CreateUpdate(aInsert, "FILE0_10H", " where doc_no = " & addstring(xDoc_No.Text))
End If
myreplaceGrd nRow
con.CommitTrans
myreplace = True
Exit Function
myerror:
prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0), , 0)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    Dim bNew As Boolean
    bNew = grid1.Row = grid1.Rows - 1
    grid1.TextMatrix(grid1.Row, 0) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    GrdDesc grid1.Row
    Grid1_AfterEdit grid1.Row, grid1.Col
    If Not bNew Then
        Unload oSearchItem
        CellPos 13, grid1.Row, grid1.Col
    End If
Else
    xDoc_No.Text = Search31.grid1.TextMatrix(Search31.grid1.Row, 0)
    CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
    If Not CardTable.EOF Then myload
    Unload Search31
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
Unload Search
End Sub
Private Sub CMD_PRINT_Click()
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
sourcetable.Open "Select * from FILE0_10 where DOC_NO = " & MyParn(xDoc_No.Text), con, adOpenForwardOnly, adLockReadOnly
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
With temptable
sourcetable.MoveFirst
Do While Not sourcetable.EOF
    ItemTable.Find " ITEM = " & MyParn(sourcetable!Item), , adSearchForward, adBookmarkFirst
    .AddNew
    !str19 = "„” ‰œ Ã—œ"
    !Date1 = xDate.Text
    !str11 = ItemTable!Item
    !str1 = ItemTable!MODELFACT0
    !str5 = ItemTable!DESCA
    !str4 = ItemTable!scal
    !str2 = ItemTable!MOSM
    !str3 = TurnValue(GetDesca("SELECT DESCA FROM FACT WHERE CODE = " & MyParn(ItemTable!Fact)), "", Null)
    !str6 = xDoc_No.Text
    !str7 = xstore.Text
    !Date1 = xDate.Text
    !val3 = sourcetable!RealBal
    !Val8 = sourcetable!Item
    !val2 = sourcetable!ComputerBal
    If sourcetable!Differ > 0 Then
        !VAL4 = Val(sourcetable!Differ)
        !val5 = Val(sourcetable!Differ * ItemTable!COST)
    Else
        !Val6 = Val(sourcetable!Differ) * -1
        !VAL7 = Val(sourcetable!Differ * ItemTable!COST) * -1
    End If
    !val1 = ItemTable!COST
    .Update
    sourcetable.MoveNext
Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\Print_Inv0.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.Report1.Action = 1
temptable.Close
Set temptable = Nothing


End Sub
Private Sub Cmd_Print2_Click()
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
cString = " SELECT  FILE0_10.doc_no, FILE0_10.item, FILE1_10.COST, FILE0_10.RealBal, FILE1_10.MODELFACT0, FILE1_10.code, FILE1_10.desca, FILE1_10.FACT, FILE1_10.SCAL, FILE1_10.COLOR, FILE1_10.C_SCAL,FILE4_10.DESCA AS CODEDESCA,FACT.DESCA AS FACTDESCA " & _
        " FROM  ((FILE0_10 INNER JOIN FILE1_10 ON FILE0_10.item = FILE1_10.ITEM) LEFT JOIN FILE4_10 ON FILE1_10.CODE = FILE4_10.CODE) LEFT JOIN FACT ON FILE1_10.FACT = FACT.CODE "
cString = cString & turn(cString) & "DOC_NO = " & MyParn(xDoc_No.Text)
        
sourcetable.Open cString, con, adOpenForwardOnly, adLockReadOnly

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
With temptable
sourcetable.MoveFirst
Do While Not sourcetable.EOF
    .AddNew
    !str7 = "ÿ»«⁄… „” ‰œ «·Ã—œ —ﬁ„ " & xDoc_No.Text
    !Date1 = xDate.Text
    !str6 = xstore.Text
    !str5 = sourcetable!code
    !str15 = TurnValue(sourcetable!CodeDesca)
    
    !str8 = sourcetable!Item
    !STR9 = TurnValue(sourcetable!FACTDESCA)
    !str1 = sourcetable!MODELFACT0
    !str2 = sourcetable!DESCA
    !str3 = sourcetable!scal
    !str4 = sourcetable!Color
    
    !VAL4 = sourcetable!RealBal
    !val5 = Val(sourcetable!RealBal & "") * Val(sourcetable!COST & "")
    .Update
    sourcetable.MoveNext
Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\Print_Inv2.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.Report1.Action = 1
temptable.Close
Set temptable = Nothing

End Sub
Private Sub cmdDelinv_Click()
    If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        con.Execute " Delete  From FILE0_100 where Doc_No = " & MyParn(xDoc_No.Text)
        con.Execute " Delete  From FILE0_10 where Doc_No = " & MyParn(xDoc_No.Text)
'       con.Execute " Delete  From FILE0_10H where Doc_No = " & MyParn(xDoc_No.Text)
        con.CommitTrans
        openCardTable
        CmdNewInv_Click
        
        Inform " „ Õ–› «·„” ‰œ »‰Ã«Õ"
    End If
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFix_Click()
Dim StockTable As New ADODB.Recordset
Dim nBal As Double
If Not MYVALIDfix Then Exit Sub
con.Execute " DELETE FROM FILE0_10 WHERE DOC_NO = " & MyParn(xDoc_No.Text)
StockTable.Open "SELECT ITEM , SUM(RealBal) AS RealBal FROM FILE0_100 WHERE DOC_NO = " & MyParn(xDoc_No.Text) & " GROUP BY ITEM ", con, adOpenStatic, adLockReadOnly, adCmdText
With StockTable
Do While Not .EOF
    nBal = RetItemBalance(!Item, xstore.BoundText, xDate.Text, con)
    SubReplaceItem !Item, !RealBal, nBal, !RealBal - nBal
   .MoveNext
Loop
'prog1.Visible = False
'myreplace

End With
End Sub
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO,DATE, CONVERT(VARCHAR(10),[DATE],111),FILE0_40.DESCA,FILE0_10H.DESCA " & _
                  " FROM (FILE0_10H INNER JOIN FILE0_40 ON FILE0_10H.Store = FILE0_40.CODE)"
If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If

Generalarray(2) = "Order by Date , DOC_NO "
Generalarray(3) = 4200
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ-«·»Ì«‰"
listarray(0, 1) = "(@@Doc_No@@6 OR " & _
                  " ##[DATE]##)"

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

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search31
Search31.Caption = "«” ⁄·«„"
Search31.Show 1
End Sub
Private Sub CmdFirst_Click()
    CardTable.MoveFirst
    myload
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

Private Sub cmdPost_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
On Error GoTo myerror
con.BeginTrans
cString = " UPDATE FILE0_10H SET FILE0_10H.closed = 1 , [ISDATE] = GETDATE() WHERE FILE0_10H.DOC_NO = " & MyParn(xDoc_No.Text)
con.Execute cString
con.CommitTrans
openCardTable
myUndo
MsgBox " „  —ÕÌ· «·„” ‰œ »‰Ã«Õ"
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
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
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
openCardTable
myUndo
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub cmdunPost_Click()
On Error GoTo myerror
con.BeginTrans
cString = " UPDATE FILE0_10H SET FILE0_10H.closed = 0 WHERE FILE0_10H.DOC_NO = " & MyParn(xDoc_No.Text)
con.Execute cString
cString = " DELETE FROM FILE0_10 WHERE FILE0_10.DOC_NO = " & MyParn(xDoc_No.Text)
con.Execute cString
con.CommitTrans
openCardTable
myUndo
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdAddItems_Click()
Dim loctable As New ADODB.Recordset, nRecordCount As Integer

cString = "SELECT FILE1_10.item,file1_10.mosm ,fact.desca as factDesca,file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color,Sum(FILE1_11.[IN] - FILE1_11.[out]) as Balance" & _
          " FROM (FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) INNER JOIN FACT ON FACT.CODE = FILE1_10.FACT"
cString = cString & turn(cString) & " FILE1_11.DATE <= " & DateSq(xDate.Text)
cString = cString & turn(cString) & " FILE1_11.STORE = " & MyParn(xstore.BoundText)

cString = cString & " GROUP BY FILE1_10.item,file1_10.mosm ,fact.desca,file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color"
cString = cString & turn(cString, " HAVING ", " AND ") & " Sum(FILE1_11.[IN] - FILE1_11.[out]) <> 0"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText


If Not (loctable.EOF And loctable.BOF) Then
    nRecordCount = loctable.RecordCount
End If
With grid1
    prog1.Visible = True
    prog1.Value = 0
    Do Until loctable.EOF
        If grid1.FindRow(loctable!Item, , 0) = -1 Then
            I = I + 1
            prog1.Value = Round(I / (nRecordCount), 2) * 100
            grid1.TextMatrix(.Rows - 1, 0) = loctable!Item
            grid1.TextMatrix(.Rows - 1, 1) = loctable!MOSM & ""
            grid1.TextMatrix(.Rows - 1, 2) = loctable!FACTDESCA
            grid1.TextMatrix(.Rows - 1, 3) = loctable!SUPP & ""
            grid1.TextMatrix(.Rows - 1, 4) = loctable!MODELFACT0 & ""
            grid1.TextMatrix(.Rows - 1, 5) = loctable!DESCA & ""
            grid1.TextMatrix(.Rows - 1, 6) = loctable!scal & ""
            grid1.TextMatrix(.Rows - 1, 7) = loctable!Color & ""
            grid1.TextMatrix(.Rows - 1, 8) = 0
            grid1.AddItem ""
        End If
        loctable.MoveNext
    Loop
    prog1.Visible = False
End With
End Sub

Private Sub Command1_Click()
For I = 1 To 160
    cDoc = RetZero(I, 6)
    con.Execute " UPDATE FILE0_10H SET CLOSED = 1 WHERE DOC_NO = " & MyParn(cDoc)
    Me.Caption = cDoc
Next I
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
bEdit = True
openCon con
'con.Execute " update file0_10h SET branch = " & addstring(cBranch) & " where substring(doc_no,1,2) = " & MyParn(cBranch)
ItemTable.Open "FILE1_10", con, adOpenStatic, adLockReadOnly, adCmdTable

cDefStore = retDef("FILE0_40")
DATA2.ConnectionString = strCon
DATA2.RecordSource = "SELECT * FROM FILE0_40"
Set xstore.RowSource = DATA2
xstore.ListField = "Desca"
xstore.BoundColumn = "Code"

If sStore <> "" Then
    xstore.BoundText = sStore
    xstore.Enabled = False
End If

Set grid1.DataSource = data1
data1.ConnectionString = strCon

openCardTable
CmdNewInv_Click


'Frame6.Visible = (cBranch = "00")
'CmdDelInv.Visible = (cBranch = "00")
'cmdSave.Visible = (cBranch = "00")
'cmdUnpost.Visible = (cBranch = "00")
'cmdPost.Visible = (cBranch = "00")
'cmdFix.Visible = (cBranch = "00")

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

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
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
    grid1.Row = grid1.Row
    grid1.Col = 8
    
    grid1.ShowCell grid1.Row, 8
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
    .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = vbWhite
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
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If Not IsNumeric(grid1.EditText) Then
        Inform "»«—ﬂÊœ «·’‰› €Ì— ”·Ì„"
        Cancel = True
        Exit Sub
    End If
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
Private Sub STOCK_MODEL_Click()
If Not MYVALID2 Then Exit Sub
StockModelfrm.Show 1
openCardTable
myUndo
End Sub
Private Sub xDate_GotFocus()
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
If xDoc_No.Text = "" Then
    If Not bIsgMsg Then Inform "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    If Not bIsMsg Then Inform "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
If xstore.BoundText = "" Then
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
If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
If xstore.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
    Exit Function
End If

MYVALID2 = True
End Function
Private Sub myload()
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xstore.BoundText = CardTable!store
xdesca.Text = CardTable!DESCA & ""
xbranch.Caption = CardTable!BRANCH & ""
xClosed.Value = IIf(CardTable!CLOSED, 1, 0)
xDoc_No.Enabled = False
myloadgrd
Handlecontrols LoadMode
CalcTotals
If Mid(xDoc_No, 1, 2) <> cBranch Then
    cmdSave.Enabled = False
    CmdDelInv.Enabled = False
End If

End Sub
Private Sub mydefine()
xDoc_No.Enabled = True
xDoc_No.Text = NewflagBranch("FILE0_10H", "DOC_NO", cBranch)
xDate.Text = Format(Date, "dd-mm-yyyy")
xstore.BoundText = cDefStore
xdesca.Text = ""
xClosed.Value = 0
xbranch.Caption = cBranch
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
cmdNewInv.Enabled = nMode = LoadMode And bEdit
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
bClosed = xClosed.Value = 1
cmdPost.Enabled = (Not bClosed) And bEdit And nMode = LoadMode
cmdUnpost.Enabled = bClosed And bEdit And nMode = LoadMode
cmdFix.Enabled = (Not bClosed) And bEdit And nMode = LoadMode
cmdSave.Enabled = (nMode = DefineMode Or Not bClosed) And bEdit
CmdDelInv.Enabled = (Not bClosed) And bEdit And nMode = LoadMod
xstore.Enabled = cDefStore = ""
STOCK_MODEL.Enabled = (Not bClosed) And bEdit
xDoc_No.Tag = nMode
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No
If xDoc_No.Text = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text)
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 0 Then
    ItemsLookupAll Me, oSearchItem
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            con.Execute "delete from FILE0_100 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
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
con.RollbackTrans
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
If Not IsNumeric(grid1.TextMatrix(Row, 0)) Then Exit Sub
nBalance = RetItemBalance(grid1.TextMatrix(Row, 0), xstore.BoundText, xDate.Text, con)
grid1.TextMatrix(Row, 9) = nBalance
Dim ItemTable As New ADODB.Recordset
Set ItemTable = ItemFind(grid1.TextMatrix(Row, 0), con)
With grid1
    If Not (ItemTable.EOF And ItemTable.BOF) Then
        grid1.TextMatrix(grid1.Row, 1) = ItemTable!MOSM
        grid1.TextMatrix(grid1.Row, 2) = GetDesca("select desca from fact where code = " & MyParn(ItemTable!Fact))
        grid1.TextMatrix(grid1.Row, 3) = ItemTable!SUPP & ""
        grid1.TextMatrix(grid1.Row, 4) = ItemTable!MODELFACT0
        grid1.TextMatrix(grid1.Row, 5) = ItemTable!DESCA
        grid1.TextMatrix(grid1.Row, 6) = ItemTable!Color
        grid1.TextMatrix(grid1.Row, 7) = ItemTable!scal
        grid1.TextMatrix(grid1.Row, 8) = IIf(grid1.TextMatrix(grid1.Row, 8) = "", 1, grid1.TextMatrix(grid1.Row, 8))
        
    End If
End With
End Sub
Private Function CalcTotals()
Dim nCount As Double
Dim nTotalQuant As Double, nTotalCost As Double
With grid1
For I = 1 To grid1.Rows - 2
    nCount = nCount + 1
'    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(I, 8))
Next
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
'.ColComboList(2) = StrList("SELECT CODE , DESCA FROM FACT ")
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
n1 = .Rows - 2
For nRow = 1 To .Rows - 1
    .TextMatrix(nRow, 9) = nRow
    n2 = n2 + Val(.TextMatrix(nRow, 8))
Next nRow
XTQ1.Text = n1
XTQ2.Text = n2

End With
End Sub
Private Sub myreplaceGrd(nRow)
Dim aInsert(2, 1)
With grid1
    For I = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
        If .TextMatrix(I, 0) <> "" Then
        
            aInsert(0, 0) = "doc_no"
            aInsert(0, 1) = addstring(xDoc_No.Text)
            
            aInsert(1, 0) = "item"
            aInsert(1, 1) = addstring(grid1.TextMatrix(I, 0))
            
            aInsert(2, 0) = "RealBal"
            aInsert(2, 1) = Val(.TextMatrix(I, 8))
            
'            aInsert(3, 0) = "ComputerBal"
'            aInsert(3, 1) = Val(.TextMatrix(I, 9))
'
'            aInsert(4, 0) = "Differ"
'            aInsert(4, 1) = Val(.TextMatrix(I, 10))
                                                                                            
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
cString = cString & turn(cString) & " DOC_NO = " & MyParn(xDoc_No.Text)
cString = cString & " ORDER BY ID"
data1.RecordSource = cString
data1.Refresh
grid1.AddItem ""
Fixgrd
End Sub
Private Function MYVALIDfix() As Boolean
If foundOther Then Exit Function
If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xstore.BoundText = "" Then
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
        If GetDesca("SELECT ITEM FROM FILE1_10 WHERE ITEM = " & MyParn(.TextMatrix(I, 0))) = "" Then
            .Select I, 0, I, 2
            MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
            Exit Function
        End If
    End If
Next
End With
MYVALIDfix = True
End Function
Private Sub openCardTable()
Set CardTable = New ADODB.Recordset
Dim cString As String
cString = "SELECT * FROM FILE0_10H"
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY DOC_NO"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    mydefine
Else
    If Trim(xDoc_No.Text) <> "" Then
        CardTable.Find "DOC_NO = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub xfilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FilterGrd grid1, xfilter.Text, 0
    End If
End Sub
Private Sub SubReplaceItem(PITEM, nReal, nComputer, nDiffer)
Dim aInsert(4, 1)
With grid1
    aInsert(0, 0) = "doc_no"
    aInsert(0, 1) = addstring(xDoc_No.Text)
        
    aInsert(1, 0) = "item"
    aInsert(1, 1) = addstring(PITEM)
        
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
myGotFocus xdesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xdesca
End Sub
Private Sub xStore_GotFocus()
myGotFocus xstore, False
End Sub
Private Sub xstore_LostFocus()
myLostFocus xstore
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
