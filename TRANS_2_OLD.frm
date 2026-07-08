VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form trans_2_OLD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕÊÌ· ’«œ— „‰ «·›—Ê⁄"
   ClientHeight    =   9990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18450
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
   ScaleHeight     =   9990
   ScaleWidth      =   18450
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmD_TXT 
      BackColor       =   &H00FFFFC0&
      Caption         =   "≈” Ì—«œ „‰ „·› TXT"
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
      Left            =   3690
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "⁄—÷"
      Top             =   45
      Width           =   1860
   End
   Begin TabDlg.SSTab sstab1 
      Height          =   5865
      Left            =   90
      TabIndex        =   58
      Top             =   3060
      Width           =   18240
      _ExtentX        =   32173
      _ExtentY        =   10345
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "«·«’‰«› «·„” ·„… »«·›⁄· »«·»«—ﬂÊœ"
      TabPicture(0)   =   "TRANS_2_OLD.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GRID2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "«·√’‰«› «· Ï  „ «—”«·Â«"
      TabPicture(1)   =   "TRANS_2_OLD.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   5325
         Left            =   -74910
         TabIndex        =   59
         Top             =   360
         Width           =   18015
         _cx             =   31776
         _cy             =   9393
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
      Begin VSFlex7Ctl.VSFlexGrid GRID2 
         Height          =   5325
         Left            =   90
         TabIndex        =   60
         Top             =   405
         Width           =   18015
         _cx             =   31776
         _cy             =   9393
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
         BackColor       =   12648447
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   12648447
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
   End
   Begin VB.CommandButton CMD_OK 
      BackColor       =   &H00FFFFC0&
      Caption         =   " „  „—«Ã⁄… «·„” ‰œ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7290
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   450
      Width           =   2445
   End
   Begin VB.CommandButton CMD_UNACCEPT 
      BackColor       =   &H000000FF&
      Caption         =   "› Õ «· ÕÊÌ·"
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
      Left            =   180
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   630
      Width           =   2085
   End
   Begin VB.Frame Frame8 
      Height          =   1410
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   1620
      Width           =   2985
      Begin VB.CommandButton CMD_sales 
         BackColor       =   &H00FFFFC0&
         Caption         =   "«·„” ‰œ ≈·Ï  ÕÊÌ· + „»Ì⁄« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   315
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   945
         Width           =   2490
      End
      Begin VB.CommandButton CMD_RETINV 
         BackColor       =   &H00FFFFC0&
         Caption         =   "«·„” ‰œ ≈·Ï „—œÊœ „‘ —Ì« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   315
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   540
         Width           =   2490
      End
      Begin VB.CommandButton CMD_TRANS 
         BackColor       =   &H00FFFFC0&
         Caption         =   "«·„” ‰œ ≈·Ï  ÕÊÌ·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   315
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   135
         Width           =   2490
      End
   End
   Begin VB.Frame Frame7 
      Height          =   600
      Left            =   3150
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   1620
      Width           =   15000
      Begin VB.TextBox xNotes 
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
         Left            =   630
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   180
         Width           =   13110
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ«  :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   13860
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   135
         Width           =   810
      End
   End
   Begin Threed.SSCommand CMD_LOOOPEN 
      Height          =   555
      Left            =   2340
      TabIndex        =   43
      Top             =   45
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   979
      _Version        =   196610
      CaptionStyle    =   1
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«” ⁄·«„  ÕÊÌ·«  €Ì— „—Õ·…"
   End
   Begin VB.CommandButton CMD_ACCEPT 
      BackColor       =   &H000080FF&
      Caption         =   "«—”«· «· ÕÊÌ· ··≈œ«—…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   180
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   945
      Width           =   2085
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Height          =   690
      Left            =   3150
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   2250
      Width           =   15000
      Begin VB.Label xTimeSend 
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
         Left            =   12195
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   225
         Width           =   2715
      End
      Begin VB.Label xTimeAccept 
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
         Left            =   4905
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   225
         Width           =   2580
      End
      Begin VB.Label xUserSend 
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
         Left            =   8685
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   225
         Width           =   2310
      End
      Begin VB.Label xUserAccept 
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
         Left            =   1350
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   225
         Width           =   2085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "User Send"
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
         Left            =   7605
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "User accept"
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
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Time Send "
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
         Left            =   11070
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Time accept"
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
         Left            =   3555
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   180
         Width           =   1245
      End
   End
   Begin VB.CommandButton TRANS_MODEL 
      Caption         =   " ÕÊÌ· „ÊœÌ·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   45
      Width           =   2085
   End
   Begin VB.Frame Frame4 
      Height          =   510
      Left            =   14850
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   8910
      Visible         =   0   'False
      Width           =   3300
      Begin VB.Label xCount 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label xTq 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   180
         Width           =   1320
      End
   End
   Begin VB.Frame Frame6 
      Height          =   645
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   8910
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   22
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
         Picture         =   "TRANS_2_OLD.frx":0038
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "TRANS_2_OLD.frx":2208
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
         TabIndex        =   23
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
         Picture         =   "TRANS_2_OLD.frx":4350
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "TRANS_2_OLD.frx":6518
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
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
         Picture         =   "TRANS_2_OLD.frx":8667
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "TRANS_2_OLD.frx":A847
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
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
         Picture         =   "TRANS_2_OLD.frx":C9A2
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "TRANS_2_OLD.frx":EB5E
      End
   End
   Begin VB.CheckBox xsent 
      Alignment       =   1  'Right Justify
      Caption         =   "Sent"
      Enabled         =   0   'False
      Height          =   195
      Left            =   11205
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   495
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CheckBox xReceived 
      Alignment       =   1  'Right Justify
      Caption         =   "xReceived"
      Enabled         =   0   'False
      Height          =   195
      Left            =   10035
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   495
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   405
      Width           =   1365
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TRANS_2_OLD.frx":10CAD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   675
         UseMaskColor    =   -1  'True
         Width           =   1275
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
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TRANS_2_OLD.frx":13226
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   12825
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   0
      Width           =   5325
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TRANS_2_OLD.frx":15589
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   510
         Left            =   1320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TRANS_2_OLD.frx":179F5
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   510
         Left            =   2655
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TRANS_2_OLD.frx":1A28F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdInform 
         Height          =   510
         Left            =   3960
         Picture         =   "TRANS_2_OLD.frx":1C83B
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame9 
      Height          =   1095
      Left            =   2295
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   540
      Width           =   3300
      Begin VB.CommandButton cmdPrint 
         Caption         =   "ÿ»«⁄… «·„” ‰œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   2025
         Picture         =   "TRANS_2_OLD.frx":1F00E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   135
         Width           =   1230
      End
      Begin VB.Label CONSTAT 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   180
         Width           =   1860
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   10035
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   630
      Width           =   8115
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
         Height          =   330
         Left            =   5400
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1500
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
         Height          =   330
         Left            =   1350
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo xStore1 
         Height          =   315
         Left            =   4140
         TabIndex        =   2
         Top             =   540
         Width           =   2760
         _ExtentX        =   4868
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
      Begin MSDataListLib.DataCombo xStore2 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   540
         Width           =   2715
         _ExtentX        =   4789
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "≈·Ì „Œ“‰ :"
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
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   495
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„‰ „Œ“‰ :"
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
         Left            =   6975
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   495
         Width           =   855
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
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   135
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
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
         Left            =   6975
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   135
         Width           =   930
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   -300
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
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
      Caption         =   "data1"
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
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   330
      Left            =   1620
      Top             =   495
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
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
      Caption         =   "data1"
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
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
      Caption         =   "data1"
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   26
      Top             =   9570
      Width           =   18450
      _ExtentX        =   32544
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
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
   Begin Threed.SSCommand cmd_excel 
      Height          =   465
      Left            =   12465
      TabIndex        =   56
      Top             =   8955
      Width           =   2370
      _ExtentX        =   4180
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
      Picture         =   "TRANS_2_OLD.frx":21438
      Caption         =   " ÕÊÌ· «ﬂ”Ì· "
      Alignment       =   4
      PictureAlignment=   1
   End
   Begin ComctlLib.ProgressBar prog1 
      Height          =   330
      Left            =   8595
      TabIndex        =   57
      Top             =   9045
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc DATA20 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
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
      Caption         =   "data1"
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
   Begin VB.Label xNotes2 
      Alignment       =   2  'Center
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
      Left            =   7290
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   90
      Width           =   5460
   End
   Begin VB.Label XDOC_TRANS 
      Alignment       =   2  'Center
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
      Left            =   7290
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   1260
      Width           =   2445
   End
   Begin VB.Label XTYPEDOC 
      Alignment       =   2  'Center
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
      Left            =   7290
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   900
      Width           =   2445
   End
   Begin VB.Label XBARNCH 
      Alignment       =   1  'Right Justify
      Caption         =   "Label9"
      Height          =   240
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   135
      Visible         =   0   'False
      Width           =   330
   End
End
Attribute VB_Name = "trans_2_OLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean, sFlagTRans As String, nFlagPublic As Double
Public sDoc_no As String
Dim bAct As Boolean
Dim cFilter As String
Dim con As New ADODB.Connection
Dim oSearchDoc As New Search3, oSearchItem As New Search3
Dim CardTable As ADODB.Recordset
Dim formMode, dDateLast As String
Const LoadMode = 0, DefineMode = 1
'5797
Private Function myreplace(Optional pRow As Long = -1, Optional bModel As Boolean) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", DateSq(xDate.Text))
aInsert = AddFlag(aInsert, "[STORE1]", addstring(xStore1.BoundText))
aInsert = AddFlag(aInsert, "[STORE2]", addstring(xStore2.BoundText))
aInsert = AddFlag(aInsert, "[isdate]", "getdate()")
aInsert = AddFlag(aInsert, "[USERSEND]", addstring(cusername))
aInsert = AddFlag(aInsert, "[branch]", addstring(cBranch))
aInsert = AddFlag(aInsert, "[NOTES]", addstring(xNotes.Text))

con.BeginTrans
'On Error GoTo myerror
If xDoc_No.Tag = DefineMode Then
    xDoc_No.Text = NewflagBranch("FILE1_61H", "DOC_NO", cBranch)
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.Text))
    con.Execute addInsert(aInsert, "FILE1_61h")
Else
    con.Execute addUpdate(aInsert, "FILE1_61h", "doc_no = " & addstring(xDoc_No.Text))
End If
If bModel Then
    myreplaceGrdModel
Else
    myreplaceGrd pRow
    myreplaceGrd2 pRow
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
'prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    bNew = grid1.Row = grid1.Rows - 1
    Dim nFound As Integer
    nFound = grid1.FindRow(oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0), , 0)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    grid1.TextMatrix(grid1.Row, 0) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 8) = "1"
    grid1_AfterEdit grid1.Row, 0
    If Not bNew Then
        CellPos 13, grid1.Row, grid1.Col
        Unload oSearchItem
    Else
        grid1.Row = grid1.Rows - 1
    End If
ElseIf ActiveControl.Name = GRID2.Name Then
    bNew = GRID2.Row = GRID2.Rows - 1
    Dim nFound As Integer
    nFound = GRID2.FindRow(oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0), , 0)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    GRID2.TextMatrix(GRID2.Row, 0) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    GRID2.TextMatrix(GRID2.Row, 8) = "1"
    grid2_AfterEdit GRID2.Row, 0
    If Not bNew Then
        CellPos 13, GRID2.Row, GRID2.Col
        Unload oSearchItem
    Else
        GRID2.Row = GRID2.Rows - 1
    End If
Else
    xDoc_No.Text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    myUndo
    Unload oSearchDoc
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
Unload Search
End Sub
Private Sub CMD_BAR_Click()
On Error GoTo myerror
    With grid1
        con.BeginTrans
        For I = 1 To .Rows - 1
            If Val(.TextMatrix(I, 8)) > 0 Then
                con.Execute "Insert Into ADDPRINT(Item,Quant,isPrint) " & _
                    " Values(" & _
                    addstring(.TextMatrix(I, 0)) & "," & _
                    addvalue(.TextMatrix(I, 8)) & "," & _
                    "1" & _
                    ")"
            End If
        Next I
        con.CommitTrans
    End With
    Inform " „ «· —ÕÌ·"
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub CMD_PrintBar_Click()
    BarCodeFRM.Show 1
End Sub
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub CMD_LOOOPEN_Click()
    LookTRansNot
End Sub
Private Sub CMD_OK_Click()
    Dim cRem As String
    If MsgBox(" „ „—«Ã⁄… «·„” ‰œ ", vbYesNo + vbDefaultButton2) = vbYes Then
        cRem = InputBox("", " »Ì«‰ „—«Ã⁄… «·„” ‰œ ")
        If nFlagPublic = 0 Then
            con.Execute "  UPDATE FILE1_61H SET Notes2 = " & addstring(cRem) & " ,  TYPEDOC = ' „  «·„—«Ã⁄…' , DOC_TRANS = " & addstring("000000") & " , ISNEW = 1 WHERE DOC_NO = " & MyParn(xDoc_No.Text)
        Else
            con.Execute "  UPDATE FR1_61H SET Notes2 = " & addstring(cRem) & " ,  TYPEDOC = ' „  «·„—«Ã⁄…' , DOC_TRANS = " & addstring("000000") & " , ISNEW = 1 WHERE DOC_NO = " & MyParn(xDoc_No.Text)
        End If
        openCardTable
        myUndo
    End If
End Sub
Private Sub CMD_sales_Click()
    Dim cDoc As String
    Dim aInsert As Variant
    If MsgBox("≈·Ï „” ‰œ  ÕÊÌ· ", vbYesNo + vbDefaultButton2) = vbYes Then
        aInsert = AddFlag(Empty, "[DATE]", DateSq(Date))
        aInsert = AddFlag(aInsert, "[STORE1]", addstring(xStore1.BoundText))
        aInsert = AddFlag(aInsert, "[STORE2]", addstring(xStore2.BoundText))
        aInsert = AddFlag(aInsert, "[isdate]", "getdate()")
        aInsert = AddFlag(aInsert, "[USERSEND]", addstring(cusername))
        aInsert = AddFlag(aInsert, "[Notes_Doc]", addstring(xNotes.Text))
        aInsert = AddFlag(aInsert, "[TRANS_DOC]", addstring(xDoc_No.Text))
        cDoc = RetZero(Newflag("FILE1_60H", "DOC_NO"))
        aInsert = AddFlag(aInsert, "DOC_NO", addstring(cDoc))
        If nFlagPublic = 0 Then
            con.Execute addInsert(aInsert, "FILE1_60h")
        Else
            con.Execute addInsert(aInsert, "FILE1_60h")
        End If
        con.Execute "  INSERT INTO FILE1_60 (DOC_NO , quant, ITEM, COST)  SELECT  " & addstring(cDoc) & " , quant, FILE1_61.ITEM, COSTITEM  From FILE1_61 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_61.ITEM WHERE FILE1_61.item <> 0 and DOC_NO = " & MyParn(xDoc_No.Text)
        con.Execute "  UPDATE FILE1_61H SET ISNEW = 1 , TYPEDOC = ' ÕÊÌ·' , DOC_TRANS = " & addstring(cDoc) & " WHERE DOC_NO = " & MyParn(xDoc_No.Text)
        con.Execute " DELETE FROM NEW1_60H WHERE DOC_NO = " & addstring(cDoc)
        For nRow = 2 To nCountBranch
            sStore = RetZero(nRow, 2)
            con.Execute "  INSERT INTO NEW1_60H (DOC_NO , BRANCH) VALUES        ( " & addstring(cDoc) & " , " & addstring(sStore) & " )  "
        Next nRow
        Inform " „  ÕÊÌ· ≈·Ï „” ‰œ  ÕÊÌ· —ﬁ„ " & cDoc
    End If
    
    If MsgBox("≈·Ï „” ‰œ „»Ì⁄«  ", vbYesNo + vbDefaultButton2) = vbYes Then
        aInsert = AddFlag(Empty, "[DATE]", DateSq(Date))
        aInsert = AddFlag(aInsert, "[mosm]", addstring(cPMosm))
        aInsert = AddFlag(aInsert, "[STORE]", addstring("000"))
        aInsert = AddFlag(aInsert, "[code]", addstring("0000"))
        aInsert = AddFlag(aInsert, "[isdate]", "getdate()")
        aInsert = AddFlag(aInsert, "[BRANCH]", addstring(cBranch))
        aInsert = AddFlag(aInsert, "[Notes]", addstring(xNotes.Text))
        aInsert = AddFlag(aInsert, "[TRANS_DOC]", addstring(xDoc_No.Text))
        cDoc = NewflagDoc(Date, xStore1.BoundText, con)
        aInsert = AddFlag(aInsert, "DOC_NO", addstring(cDoc))
        con.Execute addInsert(aInsert, "FILE6_20h")
        
        con.Execute "  INSERT INTO FILE6_20 (DOC_NO , quant, ITEM, price , man )  SELECT  " & addstring(cDoc) & " , file1_61.quant, file1_61.ITEM, file1_10.COST , " & addstring("0001") & "  From FILE1_61 inner join file1_10 on file1_61.item = file1_10.item WHERE file1_61.item <> 0 and DOC_NO = " & MyParn(xDoc_No.Text)
        con.Execute "  UPDATE FILE1_61H SET ISNEW = 1, TYPEDOC = '„»Ì⁄« ' , DOC_TRANS = " & addstring(cDoc) & " WHERE DOC_NO = " & MyParn(xDoc_No.Text)
        Inform " „  ÕÊÌ· ≈·Ï „” ‰œ „»Ì⁄«  —ﬁ„ " & cDoc
    End If
    openCardTable
    myUndo
End Sub


Private Sub CMD_UNACCEPT_Click()
On Error GoTo myerror
If MsgBox("› Õ «·„” ‰œ", vbYesNo + vbDefaultButton2) = vbYes Then
    con.Execute " UPDATE FILE1_61H  SET SENT = 0 WHERE DOC_NO = " & MyParn(xDoc_No.Text)
    openCardTable
    myUndo
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    con.BeginTrans
    On Error GoTo myerror
    con.Execute " Delete  From FILE1_61 where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute " Delete  From FILE1_62 where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute " Delete  From FILE1_61H where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    openCardTable
'    myUndo
    Inform " „ Õ–› «·„” ‰œ »‰Ã«Õ"
    CmdNewInv_Click
End If
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Sub
Private Sub cmdExit_Click()
     Unload Me
End Sub
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(1, 4)
Dim GrdArray(6, 1)

Set Generalarray(0) = Me
If nFlagPublic = 0 Then
    Generalarray(1) = "SELECT DOC_NO,DATE, CONVERT(VARCHAR(10),[DATE],111),FILE0_40.DESCA,FILE0_40_1.DESCA , TYPEDOC , Notes" & _
                      " FROM (FILE1_61H INNER JOIN FILE0_40 ON FILE1_61H.Store1 = FILE0_40.CODE) INNER JOIN FILE0_40 AS FILE0_40_1 ON FILE1_61H.STORE2 = FILE0_40_1.CODE "
Else
    Generalarray(1) = "SELECT DOC_NO,DATE, CONVERT(VARCHAR(10),[DATE],111),BRANCH_FR.DESCA,' ', TYPEDOC , Notes" & _
                      " FROM (FR1_61H INNER JOIN BRANCH_FR ON FR1_61H.Store1 = BRANCH_FR.CODE) "
End If
Generalarray(2) = "Order by Date , DOC_NO "
Generalarray(3) = 4200
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ"
listarray(0, 1) = "(@@Doc_No@@6 OR " & _
                  " ##[DATE]##)"

listarray(1, 0) = "„·«ÕŸ«  "
listarray(1, 1) = "%%Notes%%"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "„‰ „Œ“‰"
GrdArray(3, 1) = 2000

GrdArray(4, 0) = "≈·Ì „Œ“‰"
If nFlagPublic = 1 Then
    GrdArray(4, 1) = 2000
Else
    GrdArray(4, 1) = 0
End If

GrdArray(5, 0) = " "
GrdArray(5, 1) = 1000

GrdArray(6, 0) = "„·«ÕŸ«  "
GrdArray(6, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
Load oSearchDoc
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
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
xStore1.SetFocus
Err.Clear
End Sub
Private Sub cmdPrint_Click()
    doprint
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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = 83 Then cmdSave_Click
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Me.Caption = sCaption
openCon con

CmdDelInv.Visible = (cBranch <> "00")
cmdSave.Visible = (cBranch <> "00")
TRANS_MODEL.Visible = (cBranch <> "00")

CMD_UNACCEPT.Visible = (bSupermode And cBranch > "00")
Frame8.Visible = (cBranch = "00")
Set CardTable = New ADODB.Recordset
data1.ConnectionString = strCon

CMD_TRANS.Visible = (nFlagPublic = 0)
CMD_sales.Visible = (nFlagPublic = 0)

If nFlagPublic = 0 Then
    data1.ConnectionString = strCon
    data1.RecordSource = "SELECT * FROM FILE0_40"
    Set xStore1.RowSource = data1
    xStore1.ListField = "Desca"
    xStore1.BoundColumn = "Code"
    xStore1.BoundText = cBranchStore

    xStore1.Enabled = False
    xStore2.Enabled = False
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "SELECT * FROM FILE0_40"
    Set xStore2.RowSource = DATA2
    xStore2.ListField = "Desca"
    xStore2.BoundColumn = "Code"
Else
    data1.ConnectionString = strCon
    data1.RecordSource = "SELECT * FROM BRANCH_FR"
    Set xStore1.RowSource = data1
    xStore1.ListField = "Desca"
    xStore1.BoundColumn = "Code"
    xStore1.BoundText = cBranchStore
    xStore1.Enabled = False
    xStore2.Visible = False
End If
Set grid1.DataSource = DATA10
DATA10.ConnectionString = strCon

Set GRID2.DataSource = DATA20
DATA20.ConnectionString = strCon

openCardTable
myUndo
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Unload oSearchItem
Unload oSearchDoc
If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
Err.Clear
Set Transfrm = Nothing
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = 0 Then GrdDesc Row
CalcTotals
If Not validRow(Row) Then Exit Sub

With grid1
If myreplace(Row) Then
    If xDoc_No.Tag = DefineMode Then xDoc_No.Tag = LoadMode
Else
    bNew = False
End If
myloadgrd
myloadgrd2
If bNew Then
    grid1.Row = grid1.Row
    grid1.Col = 8
    
    grid1.ShowCell grid1.Row, 8
End If
End With



'If Row = grid1.Rows - 1 Then myaddItem
End Sub

Private Sub grid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = 0 Then GrdDesc2 Row
'CalcTotals2
If Not validRow(Row) Then Exit Sub

With GRID2
If myreplace(Row) Then
    If xDoc_No.Tag = DefineMode Then xDoc_No.Tag = LoadMode
Else
    bNew = False
End If
myloadgrd2
If bNew Then
    GRID2.Row = GRID2.Row
    GRID2.Col = 8
    
    GRID2.ShowCell GRID2.Row, 8
End If
End With



'If Row = grid1.Rows - 1 Then myaddItem
End Sub

Private Sub grid1_EnterCell()
With grid1
If (grid1.Col = 0 Or grid1.Col = 8 Or grid1.Col = 10) And xsent.Value = 0 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End With
End Sub
Private Sub grid2_EnterCell()
With GRID2
If (.Col = 0 Or .Col = 8 Or .Col = 10) And (xsent.Value = 0) And cBranch = "00" And (XTYPEDOC.Caption = "") Then
    .Editable = flexEDKbdMouse
Else
    .Editable = flexEDNone
End If
End With
End Sub
Private Sub Grid1_GotFocus()
With grid1
    If grid1.Row <= 1 And grid1.Rows > 1 Then
    .Select 1, 0, 1, 0
    '.ShowCell 1, 0
    End If
End With
End Sub
Private Sub Grid2_GotFocus()
With GRID2
    If GRID2.Row <= 1 And GRID2.Rows > 1 Then
    .Select 1, 0, 1, 0
    End If
End With
End Sub

Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 2 And Trim(grid1.TextMatrix(Row, 3)) = "" Then
    nBalance = LastBalance(grid1.TextMatrix(Row, 0), xStore1.BoundText, con)
    grid1.TextMatrix(Row, 9) = nBalance
End If
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 And (grid1.EditText) <> "" Then
    With grid1
    If .EditText <> "" Then
        If UCase(Mid(.EditText, 1, 1)) = "J" Then
            cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE2 = " & MyParn(.EditText))
        ElseIf UCase(Mid(.EditText, 1, 1)) = "S" Then
            cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE2 = " & MyParn(.EditText))
        Else
            cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE = " & MyParn(.EditText))
        End If
        If cMyItem <> "" Then
            .EditText = Val(cMyItem)
        End If
    End If
    End With
End If

If Col = 0 And IsNumeric(grid1.EditText) Then
    If Not IsNumeric(grid1.EditText) Then
        Inform "»«—ﬂÊœ «·’‰› €Ì— ”·Ì„"
        Cancel = True
        Exit Sub
    End If
    
    If Not validItem(grid1.EditText, con) Then
        Inform "»«—ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
        Cancel = True
    End If
    nFound = FoundOtheritem(Row, 0, Trim(grid1.EditText))
    If nFound <> -1 Then
        Inform "«·’‰› „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & nFound
        Dim sInput As String
        sInput = InputBox("«÷«›… ⁄œœ", "«÷«›…", "1")
        If IsNumeric(sInput) Then
            grid1.TextMatrix(nFound, 8) = Val(grid1.TextMatrix(nFound, 8)) + Val(sInput)
            myreplaceGrd (nFound)
            grid1.EditText = ""
        End If
        Cancel = True
    End If
End If
End Sub

Private Sub Grid2_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 And (GRID2.EditText) <> "" Then
    With GRID2
    If .EditText <> "" Then
        If UCase(Mid(.EditText, 1, 1)) = "J" Then
            cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE2 = " & MyParn(.EditText))
        ElseIf UCase(Mid(.EditText, 1, 1)) = "S" Then
            cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE2 = " & MyParn(.EditText))
        Else
            cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE = " & MyParn(.EditText))
        End If
        If cMyItem <> "" Then
            .EditText = Val(cMyItem)
        End If
    End If
    End With
End If

If Col = 0 And IsNumeric(.EditText) Then
    If Not IsNumeric(.EditText) Then
        Inform "»«—ﬂÊœ «·’‰› €Ì— ”·Ì„"
        Cancel = True
        Exit Sub
    End If
    
    If Not validItem(.EditText, con) Then
        Inform "»«—ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
        Cancel = True
    End If
    nFound = FoundOtheritem(Row, 0, Trim(.EditText))
    If nFound <> -1 Then
        Inform "«·’‰› „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & nFound
        Dim sInput As String
        sInput = InputBox("«÷«›… ⁄œœ", "«÷«›…", "1")
        If IsNumeric(sInput) Then
            .TextMatrix(nFound, 8) = Val(.TextMatrix(nFound, 8)) + Val(sInput)
            myreplaceGrd2 (nFound)
            GRID2.EditText = ""
        End If
        Cancel = True
    End If
End If
End Sub
Private Sub TRANS_MODEL_Click()
If grid1.Row = 0 Then grid1.Row = grid1.Rows - 1
Set transModelFrm.myForm = Me
transModelFrm.nColItem = 0
transModelFrm.nColQuant = 8
transModelFrm.strItem = grid1.TextMatrix(grid1.Row, 0)
transModelFrm.strStore = xStore1.BoundText
transModelFrm.bshowBal = True
transModelFrm.Show 1
End Sub
Private Sub xDate_GotFocus()
xDate.SelStart = 0
xDate.SelLength = Len(xDate.Text)
End Sub
Private Sub xDoc_No_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CmdInform_Click
End Sub
Private Function MYVALID(Optional bCheckRows As Boolean = True) As Boolean
If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xStore1.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ «·«Ê·"
    Exit Function
End If

'If xStore2.BoundText = "" Then
'    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ «·À«‰Ì"
'    Exit Function
'End If

If bCheckRows Then
    If grid1.Rows < 3 Then
        MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
        Exit Function
    End If

    With grid1
    For I = 1 To .Rows - 2
        If .TextMatrix(I, 0) = "" Then
            .Select I, 0, I, grid1.Cols - 1
            MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
            Exit Function
        ElseIf Not validItem(.TextMatrix(I, 0), con) Then
            MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
            Exit Function
        End If
        If Val(.TextMatrix(I, 8)) = 0 Then
            .Select I, 0, I, grid1.Cols - 1
            MsgBox "ﬂ„Ì… «·’‰› €Ì— „”Ã·…"
            Exit Function
        End If
    Next
    End With
End If
MYVALID = True
End Function

Private Function MYVALID2() As Boolean
If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
If xStore1.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ «·«Ê·"
    Exit Function
End If

If xStore2.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ «·À«‰Ì"
    Exit Function
End If
MYVALID2 = True
End Function
Private Sub myload()
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xStore1.BoundText = CardTable!STORE1
xStore2.BoundText = CardTable!STORE2 & ""
xNotes.Text = CardTable!NOTES & ""
xNotes2.Caption = CardTable!NOTES2 & ""

XTYPEDOC.Caption = CardTable!typedoc & ""
XDOC_TRANS.Caption = CardTable!DOC_TRANS & ""

CMD_TRANS.Enabled = (XTYPEDOC.Caption = "")
CMD_RETINV.Enabled = (XTYPEDOC.Caption = "")
CMD_sales.Enabled = (XTYPEDOC.Caption = "")
CMD_OK.Enabled = (XTYPEDOC.Caption = "")


xUserAccept.Caption = CardTable!UserAccept & ""
xUserSend.Caption = CardTable!USERSEND & ""
xTimeAccept.Caption = CardTable!TimeAccept & ""
xTimeSend.Caption = CardTable!TimeSend & ""
xReceived.Value = IIf(CardTable!Received, 1, 0)
xsent.Value = IIf(CardTable!SENT, 1, 0)

xDoc_No.Enabled = False
xDoc_No.Tag = LoadMode
myloadgrd
myloadgrd2

Handlecontrols LoadMode
StatusBar1.Panels(1).Text = "”Ã· " & CardTable.AbsolutePosition & " „‰ " & CardTable.RecordCount
StatusBar1.Panels(2).Text = "⁄œœ «·”Ã·«  : " & CardTable.RecordCount
grid1.Select grid1.Rows - 1, 0
grid1.ShowCell grid1.Rows - 1, 0
CalcTotals
End Sub
Private Sub mydefine()
xDoc_No.Enabled = True
xDoc_No.Text = NewflagBranch("FILE1_61H", "DOC_NO", cBranch)
xDate.Text = Format(Date, "dd-mm-yyyy")
xStore1.BoundText = cBranchStore
xStore2.BoundText = "000"
StatusBar1.Panels(1).Text = ""
xCount.Caption = ""
xNotes.Text = ""
xNotes2.Caption = ""
XTYPEDOC.Caption = ""
XDOC_TRANS.Caption = ""

CMD_TRANS.Enabled = True
CMD_RETINV.Enabled = True
CMD_sales.Enabled = True
CMD_OK.Enabled = True


xTq.Caption = ""
xsent.Value = 0
StatusBar1.Panels(3).Text = ""
StatusBar1.Panels(4).Text = ""

xUserAccept.Caption = ""
xUserSend.Caption = ""
xTimeAccept.Caption = ""
xTimeSend.Caption = ""


If cmdSave.Enabled Then
    StatusBar1.Panels(1).Text = "«÷«›… ”Ã· —ﬁ„ " & (CardTable.RecordCount + 1)
    StatusBar1.Panels(2).Text = "⁄œœ «·”Ã·«  : " & CardTable.RecordCount
ElseIf CardTable.EOF And CardTable.BOF Then
    StatusBar1.Panels(1).Text = ""
    StatusBar1.Panels(2).Text = "·«  ÊÃœ ”Ã·« "
End If
grid1.Rows = 1
GRID2.Rows = 1
myaddItem
myaddItem2

Handlecontrols DefineMode
Fixgrd
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = nMode = LoadMode
cmdSave.Enabled = (xsent.Value = 0)
CmdDelInv.Enabled = (nMode = LoadMode) And (xsent.Value = 0)
CmdDelInv.Enabled = nMode = LoadMode And (xsent.Value = 0)
TRANS_MODEL.Enabled = nMode = LoadMode And (xsent.Value = 0)

CMD_ACCEPT.Enabled = nMode = LoadMode And (xsent.Value = 0)

cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2 And sDoc_no = ""
xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
End Sub
Private Sub xDoc_No_LostFocus()
If xDoc_No.Text = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text, 6)
If CardTable.BOF And CardTable.BOF Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 0 And xsent.Value = 0 Then
    ItemsLookupAll Me, oSearchItem
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And nFlag = 0 And xsent.Value = 0 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            con.Execute "delete from file1_61 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
        CalcTotals
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub Grid2_KeyUp(KeyCode As Integer, Shift As Integer)
With GRID2

If KeyCode = 112 And .Col = 0 And (XTYPEDOC.Caption = "") And cBranch = "00" Then
    ItemsLookupAll Me, oSearchItem
ElseIf KeyCode = 13 Then
    CellPos2 KeyCode, .Row, .Col
ElseIf KeyCode = 46 And .Row <> .Rows - 1 And nFlag = 1 And (XTYPEDOC.Caption = "") And cBranch = "00" Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        If .TextMatrix(.Row, .Cols - 1) <> "" Then
            con.Execute "delete from file1_62 where ID = " & .TextMatrix(.Row, .Cols - 1)
        End If
        .RemoveItem .Row
'        CalcTotals2
    End If
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub

Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
If grid1.TextMatrix(Row, 0) = "" Then Exit Sub
If grid1.TextMatrix(Row, 1) = "" Then Exit Sub
KeyCode = 0
If Col = 0 Then
    If IsNumeric(grid1.TextMatrix(grid1.Row, 0)) Then grid1.Col = 8
ElseIf Col = 8 Then
    If Row <= grid1.Rows - 2 Then
        grid1.ShowCell grid1.Row, 0
        grid1.Row = Row + 1
        grid1.Col = 0
    Else
        grid1.Col = 8
    End If
End If
End Sub

Private Sub CellPos2(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
With GRID2
If .TextMatrix(Row, 0) = "" Then Exit Sub
If .TextMatrix(Row, 1) = "" Then Exit Sub
KeyCode = 0
If Col = 0 Then
    If IsNumeric(.TextMatrix(.Row, 0)) Then .Col = 8
ElseIf Col = 8 Then
    If Row <= .Rows - 2 Then
        .ShowCell .Row, 0
        .Row = Row + 1
        .Col = 0
    Else
        .Col = 8
    End If
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
Select Case Col
    Case 0
        If KeyCode = 27 Then Exit Sub
        If KeyCode = 112 Then ItemsLookupAll Me, oSearchItem
End Select
End Sub
Private Sub grid2_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos2 KeyCode, Row, Col
Select Case Col
    Case 0
        If KeyCode = 27 Then Exit Sub
        If KeyCode = 112 Then ItemsLookupAll Me, oSearchItem
End Select
End Sub

Private Sub GrdDesc(Row)
Dim nBalance As Double
If Not IsNumeric(grid1.TextMatrix(Row, 0)) Then Exit Sub
nBalance = LastBalance(grid1.TextMatrix(Row, 0), xStore1.BoundText, con)
grid1.TextMatrix(Row, 9) = nBalance

Dim ItemTable As New ADODB.Recordset
Set ItemTable = ItemFind(grid1.TextMatrix(Row, 0), con)
With grid1
    If Not ItemTable.EOF Then
        If ItemTable!code <> .TextMatrix(1, 3) And .Row > 1 Then
            MsgBox " „Ê—œ „Œ ·› ⁄‰ „Ê—œ «·„” ‰œ "
            grid1.TextMatrix(grid1.Row, 0) = ""
            Exit Sub
        End If
        grid1.TextMatrix(grid1.Row, 1) = ItemTable!MOSM
        grid1.TextMatrix(grid1.Row, 2) = GetField("select desca from fact where code = " & MyParn(ItemTable!Fact)) & ""
        grid1.TextMatrix(grid1.Row, 3) = ItemTable!code & ""
        grid1.TextMatrix(grid1.Row, 4) = ItemTable!MODELFACT0
        grid1.TextMatrix(grid1.Row, 5) = ItemTable!DESCA
        grid1.TextMatrix(grid1.Row, 6) = ItemTable!SCAL
        grid1.TextMatrix(grid1.Row, 7) = ItemTable!Color
    End If
End With
Set ItemTable = Nothing
If Val(grid1.TextMatrix(Row, 9)) <= 0 Then
'    Inform "·« ÌÊÃœ —’Ìœ ··’‰›"
End If
End Sub
Private Sub GrdDesc2(Row)
Dim nBalance As Double
With GRID2
If Not IsNumeric(.TextMatrix(Row, 0)) Then Exit Sub

Dim ItemTable As New ADODB.Recordset
Set ItemTable = ItemFind(.TextMatrix(Row, 0), con)

    If Not ItemTable.EOF Then
        If ItemTable!code <> .TextMatrix(1, 3) And .Row > 1 Then
            MsgBox " „Ê—œ „Œ ·› ⁄‰ „Ê—œ «·„” ‰œ "
            .TextMatrix(.Row, 0) = ""
            Exit Sub
        End If
        .TextMatrix(.Row, 1) = ItemTable!MOSM
        .TextMatrix(.Row, 2) = GetField("select desca from fact where code = " & MyParn(ItemTable!Fact)) & ""
        .TextMatrix(.Row, 3) = ItemTable!code & ""
        .TextMatrix(.Row, 4) = ItemTable!MODELFACT0
        .TextMatrix(.Row, 5) = ItemTable!DESCA
        .TextMatrix(.Row, 6) = ItemTable!SCAL
        .TextMatrix(.Row, 7) = ItemTable!Color
    End If
End With
Set ItemTable = Nothing
End Sub

Private Function CalcTotals()
Dim nTotalQuant As Double
With grid1
For I = 1 To grid1.Rows - 2
    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(I, 8))
Next
StatusBar1.Panels(3) = ""
StatusBar1.Panels(4) = ""
If grid1.Rows > 1 Then StatusBar1.Panels(3) = "⁄œœ «·√’‰«› : " & grid1.Rows - 1
If nTotalQuant > 0 Then StatusBar1.Panels(4) = "≈Ã„«·Ì «·ﬂ„Ì… : " & nTotalQuant
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
Private Sub foundOther()
Dim nRow As Long
For I = 1 To grid1.Rows - 2
    nRow = grid1.FindRow(grid1.TextMatrix(I, 0), I + 1, 0)
    If nRow > -1 Then
        Inform "«·’‰› " & grid1.TextMatrix(nRow, 1) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ ====> " & nRow
    End If
Next
End Sub
Private Sub doprint()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim nTot As Double
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
If nFlagPublic = 0 Then
    nTot = Val(GetDesca("SELECT SUM(QUANT) FROM FILE1_61 WHERE DOC_NO = " & MyParn(xDoc_No.Text)) & "")
Else
    nTot = Val(GetDesca("SELECT SUM(QUANT) FROM FR1_61 WHERE DOC_NO = " & MyParn(xDoc_No.Text)) & "")
End If
For I = 1 To grid1.Rows - 2
    temptable.AddNew
    temptable!str6 = xDoc_No.Text
    temptable!str8 = TurnValue(xNotes.Text)
    temptable!Date1 = DateValue(xDate.Text)
    temptable!str14 = TurnValue(xStore1.Text)
    temptable!str15 = TurnValue(xStore2.Text)
    
    temptable!str4 = TurnValue(grid1.TextMatrix(I, 0))
    temptable!str16 = GetDesca("SELECT MODEL FROM FILE1_10 WHERE ITEM = " & MyParn(grid1.TextMatrix(I, 0)))
    temptable!STR11 = TurnValue(grid1.TextMatrix(I, 1))
    temptable!str12 = TurnValue(grid1.TextMatrix(I, 4))
    temptable!str10 = TurnValue(grid1.TextMatrix(I, 2))
    temptable!str2 = TurnValue(grid1.TextMatrix(I, 5))
    temptable!str5 = TurnValue(grid1.TextMatrix(I, 6))
    temptable!STR7 = TurnValue(grid1.TextMatrix(I, 7))
    temptable!val1 = Val(grid1.TextMatrix(I, 8))
    temptable!val2 = Val(GetDesca("select price from file1_10 where item = " & grid1.TextMatrix(I, 0)) & "")
    temptable!val18 = nTot
    temptable!str1 = TurnValue(grid1.TextMatrix(I, 2)) & TurnValue(grid1.TextMatrix(I, 1)) & RetZero(TurnValue(grid1.TextMatrix(I, 4)), 10)
    temptable!Val10 = I
    temptable.Update
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\TRANS.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
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
With grid1
'                   0            1      2               3        4           5       6           7          8           9           10
.FormatString = "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "«·ﬂ„Ì…|" & "«·—’Ìœ|" & "„·«ÕŸ« |"
.WordWrap = True
.RowHeight(0) = 500
.ColWidth(0) = 1000
.ColWidth(1) = 1000
.ColWidth(2) = 2000
.ColWidth(3) = 700
.ColWidth(4) = 1500
.ColWidth(5) = 3600
.ColWidth(6) = 900
.ColWidth(7) = 900
.ColWidth(8) = 900
.ColWidth(9) = 900
.ColWidth(10) = 4000

.ColHidden(.Cols - 1) = True
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
End With
End Sub
Private Sub Fixgrd2()
With GRID2
'                   0            1      2               3        4           5       6           7          8           9           10
.FormatString = "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "«·ﬂ„Ì…|" & "«·—’Ìœ|" & "„·«ÕŸ« |"
.WordWrap = True
.RowHeight(0) = 500
.ColWidth(0) = 1000
.ColWidth(1) = 1000
.ColWidth(2) = 2000
.ColWidth(3) = 700
.ColWidth(4) = 1500
.ColWidth(5) = 3600
.ColWidth(6) = 900
.ColWidth(7) = 900
.ColWidth(8) = 900
.ColWidth(9) = 900
.ColWidth(10) = 4000

.ColHidden(.Cols - 1) = True
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
End With
End Sub
Private Sub myreplaceGrd(nRow)
Dim aInsert As Variant, I As Long, bSaveRow As Boolean
With grid1
    For I = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
        aInsert = AddFlag(Empty, "[DOC_NO]", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "[ITEM]", addvalue(grid1.TextMatrix(I, 0)))
        aInsert = AddFlag(aInsert, "[SUBNOTES]", addstring(grid1.TextMatrix(I, 10)))
        aInsert = AddFlag(aInsert, "[QUANT]", Val(grid1.TextMatrix(I, 8)))
        aInsert = AddFlag(aInsert, "[COST]", Val(ItemField(grid1.TextMatrix(I, 0), "COST", con) & ""))
        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE1_61")
        Else
            con.Execute addUpdate(aInsert, "FILE1_61", "ID = " & grid1.TextMatrix(I, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub myreplaceGrd2(nRow)
Dim aInsert As Variant, I As Long, bSaveRow As Boolean
With GRID2
    For I = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
        aInsert = AddFlag(Empty, "[DOC_NO]", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "[ITEM]", addvalue(grid1.TextMatrix(I, 0)))
        aInsert = AddFlag(aInsert, "[SUBNOTES]", addstring(grid1.TextMatrix(I, 10)))
        aInsert = AddFlag(aInsert, "[QUANT]", Val(grid1.TextMatrix(I, 8)))
        aInsert = AddFlag(aInsert, "[COST]", Val(ItemField(grid1.TextMatrix(I, 0), "COST", con) & ""))
        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE1_62")
        Else
            con.Execute addUpdate(aInsert, "FILE1_62", "ID = " & grid1.TextMatrix(I, .Cols - 1))
        End If
    Next
End With
End Sub

Private Sub myloadgrd()
'                    0               1               2               3           4                       5               6               7               8                   9
If nFlagPublic = 0 Then
    cString = "SELECT FILE1_61.item , file1_10.mosm , fact.desca , file1_10.code , file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color ,  FILE1_61.quant , ' ' as bal , FILE1_61.SubNotes ,FILE1_61.ID from (FILE1_61 inner join file1_10 on file1_10.item = FILE1_61.item ) INNER join fact on fact.code = file1_10.fact "
Else
    cString = "SELECT FR1_61.item , file1_10.mosm , fact.desca , file1_10.code , file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color ,  FR1_61.quant , ' ' as bal , FR1_61.SubNotes ,FR1_61.ID from (FR1_61 inner join file1_10 on file1_10.item = FR1_61.item ) INNER join fact on fact.code = file1_10.fact "
End If
cString = cString & turn(cString) & "DOC_NO = " & MyParn(xDoc_No.Text)
cString = cString & " ORDER BY ID"
DATA10.RecordSource = cString
DATA10.Refresh
myaddItem
Fixgrd
End Sub
Private Sub myloadgrd2()
'                    0               1               2               3           4                       5               6               7               8                   9
If nFlagPublic = 0 Then
    cString = "SELECT FILE1_62.item , file1_10.mosm , fact.desca , file1_10.code , file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color ,  FILE1_62.quant , ' ' as bal , FILE1_62.SubNotes ,FILE1_62.ID from (FILE1_62 inner join file1_10 on file1_10.item = FILE1_62.item ) INNER join fact on fact.code = file1_10.fact "
Else
    cString = "SELECT FR1_62.item , file1_10.mosm , fact.desca , file1_10.code , file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color ,  FR1_62.quant , ' ' as bal , FR1_62.SubNotes ,FR1_62.ID from (FR1_62 inner join file1_10 on file1_10.item = FR1_62.item ) INNER join fact on fact.code = file1_10.fact "
End If
cString = cString & turn(cString) & "DOC_NO = " & MyParn(xDoc_No.Text)
cString = cString & " ORDER BY ID"
DATA20.RecordSource = cString
DATA20.Refresh
myaddItem2
Fixgrd2
End Sub

Private Sub openCardTable()
Dim cBranchStore As String
cBranchStore = retFlag(aPassword, "store")
Set CardTable = New ADODB.Recordset
Dim cString As String
cFilter = ""
If nFlagPublic = 0 Then
    cString = "SELECT * FROM FILE1_61H"
Else
    cString = "SELECT * FROM FR1_61H"
End If
If sDoc_no <> "" Then cString = cString & " WHERE DOC_NO = " & MyParn(sDoc_no)

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
Sub myProcAddModel()
If Not MYVALID(False) Then Exit Sub
If myreplace(, True) Then
    Inform " „  «÷«›… «·„ÊœÌ·«  »‰Ã«Õ"
End If
openCardTable
myUndo

'myloadgrd
grid1.ShowCell grid1.Rows - 1, 2 + 1
grid1.Select grid1.Rows - 1, 2 + 1
End Sub
Private Sub xStore2_Validate(Cancel As Boolean)
    If xStore2.BoundText = xStore1.BoundText Then
        xStore2.BoundText = ""
        Cancel = True
    End If
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow Then
    If OldRow <> .Rows - 1 And OldRow <> 0 And .TextMatrix(OldRow, .Cols - 1) = "" Then
        If Not validRow(OldRow) Then
            .RemoveItem OldRow
            CalcTotals
        End If
    End If
End If
End With
End Sub
Private Sub grid2_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With GRID2
If OldRow <> NewRow Then
    If OldRow <> .Rows - 1 And OldRow <> 0 And .TextMatrix(OldRow, .Cols - 1) = "" Then
        If Not validRow(OldRow) Then
            .RemoveItem OldRow
            CalcTotals
        End If
    End If
End If
End With
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
If Not validRow(grid1.Row) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    grid1.RemoveItem grid1.Row
    CalcTotals
End If
End Sub
Private Sub Grid2_Validate(Cancel As Boolean)
If Not validRow(GRID2.Row) And GRID2.Row <> GRID2.Rows - 1 And GRID2.Row <> 0 And GRID2.TextMatrix(GRID2.Row, GRID2.Cols - 1) = "" Then
    GRID2.RemoveItem GRID2.Row
'    CalcTotals2
End If
End Sub
Private Function validRow(Row) As Boolean
With grid1
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
If grid1.TextMatrix(Row, 8) = "" Then Exit Function
End With
validRow = True
End Function
Private Function validRow2(Row) As Boolean
With GRID2
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
If grid1.TextMatrix(Row, 8) = "" Then Exit Function
End With
validRow2 = True
End Function
Private Sub myaddItem()
    If nFlag <> 1 Then grid1.AddItem ""
End Sub
Private Sub myaddItem2()
    If nFlag <> 1 Then GRID2.AddItem ""
End Sub
Private Function myreplaceGrdModel() As Boolean
Dim cString As String, aitem As Variant
cString = " DELETE FROM FILE1_61 FROM FILE1_61 INNER JOIN FILE1_10 ON FILE1_61.ITEM = FILE1_10.ITEM"
cString = cString & turn(cString) & "DOC_NO = " & MyParn(xDoc_No.Text)
cString = cString & turn(cString) & "FILE1_10.MODEL = " & MyParn(transModelFrm.xModel.Text)
con.Execute cString, nDel
With transModelFrm.grid1
    For nRow = 3 To .Rows - 1
        For nCol = 2 To .Cols - 1
            If Val(.TextMatrix(nRow, nCol)) <> 0 Then
                aitem = ItemFields(transModelFrm.GRID2.TextMatrix(nRow, nCol), con)
                If Not IsEmpty(aitem) Then
                    cString = "Insert into FILE1_61 (doc_no,item,cost,Quant)" & _
                               "Values(" & _
                               addstring(xDoc_No.Text) & "," & _
                               addvalue(retFlag(aitem, "item")) & "," & _
                               Val(retFlag(aitem, "cost") & "") & "," & _
                               Val(.TextMatrix(nRow, nCol)) & _
                               ")"
                    con.Execute cString
                End If
            End If
        Next
    Next
End With
End Function
Private Sub CMD_ACCEPT_Click()
    On Error GoTo myerror
If MsgBox("«—”«· «· ÕÊÌ· ··≈œ«—… ", vbYesNo + vbDefaultButton2) = vbYes Then
    strConShop = LoadConStringshop
    If openConShop(conShop) <> "ok" Then
        MsgBox "·« ÌÊÃœ « ’«· »«·”Ì—›—"
        Exit Sub
    End If
    CONSTAT.Visible = True
    If SendTransHader Then
        SendTransHaderSub
    End If
    con.Execute " UPDATE FILE1_61H  SET SENT = 1 , TIMESEND = GETDATE() , USERSEND = " & addstring(cusername) & " WHERE DOC_NO = " & MyParn(xDoc_No.Text)
    Inform " „ «—”«· «· ÕÊÌ· «·≈œ«—…"
    openCardTable
    myUndo

End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function SendTransHader()
On ERRROR GoTo myerror
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", DateSq(xDate.Text))
aInsert = AddFlag(aInsert, "[STORE1]", addstring(xStore1.BoundText))
aInsert = AddFlag(aInsert, "[STORE2]", addstring(xStore2.BoundText))
aInsert = AddFlag(aInsert, "[BRANCH]", addstring(cBranch))
aInsert = AddFlag(aInsert, "[notes]", addstring(xNotes.Text))
aInsert = AddFlag(aInsert, "[USERSEND]", addstring(cusername))
aInsert = AddFlag(aInsert, "[TIMESEND]", "GETDATE()")
aInsert = AddFlag(aInsert, "[SENT]", 1)
aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.Text))
If cBranch < "60" Then
    conShop.Execute " DELETE FROM FILE1_61  WHERE DOC_NO = " & MyParn(xDoc_No.Text)
    conShop.Execute " DELETE FROM FILE1_61h WHERE DOC_NO = " & MyParn(xDoc_No.Text)
    conShop.Execute addInsert(aInsert, "FILE1_61h")
Else
    conShop.Execute " DELETE FROM FR1_61     WHERE DOC_NO = " & MyParn(xDoc_No.Text)
    conShop.Execute " DELETE FROM FR1_61h   WHERE DOC_NO = " & MyParn(xDoc_No.Text)
    conShop.Execute addInsert(aInsert, "FR1_61h")
End If
SendTransHader = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function SendTransHaderSub()
On ERRROR GoTo myerror

Dim aInsert As Variant, cStr1 As String
With grid1
    For nRow = 1 To .Rows - 1
        aInsert = AddFlag(Empty, "[DOC_NO]", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "[ITEM]", Val(.TextMatrix(nRow, 0)))
        aInsert = AddFlag(aInsert, "[subnotes]", addstring(.TextMatrix(nRow, 10)))
        aInsert = AddFlag(aInsert, "[QUANT]", Val(.TextMatrix(nRow, 8)))
        If cBranch < "60" Then
            cStr1 = cStr1 & addInsert(aInsert, "FILE1_61") & " ; "
        Else
            cStr1 = cStr1 & addInsert(aInsert, "FR1_61") & " ; "
        End If
    Next nRow
    conShop.Execute cStr1
    SendTransHaderSub = True
End With
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Sub CMD_TRANS_Click()
    Dim cDoc As String
    Dim aInsert As Variant
    If MsgBox("≈·Ï „” ‰œ  ÕÊÌ· ", vbYesNo + vbDefaultButton2) = vbYes Then
        aInsert = AddFlag(Empty, "[DATE]", DateSq(Date))
        aInsert = AddFlag(aInsert, "[STORE1]", addstring(xStore1.BoundText))
        aInsert = AddFlag(aInsert, "[STORE2]", addstring(xStore2.BoundText))
        aInsert = AddFlag(aInsert, "[isdate]", "getdate()")
        aInsert = AddFlag(aInsert, "[USERSEND]", addstring(cusername))
        aInsert = AddFlag(aInsert, "[Notes_Doc]", addstring(xNotes.Text))
        aInsert = AddFlag(aInsert, "[TRANS_DOC]", addstring(xDoc_No.Text))
        
        cDoc = RetZero(Newflag("FILE1_60H", "DOC_NO"))
        aInsert = AddFlag(aInsert, "DOC_NO", addstring(cDoc))
        con.Execute addInsert(aInsert, "FILE1_60h")
        
        con.Execute "  INSERT INTO FILE1_60 (DOC_NO , quant, ITEM, COST)  SELECT  " & addstring(cDoc) & " , quant, FILE1_61.ITEM, COSTITEM  From FILE1_61 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_61.ITEM WHERE FILE1_61.item <> 0 and DOC_NO = " & MyParn(xDoc_No.Text)
        con.Execute "  UPDATE FILE1_61H SET ISNEW = 1 , TYPEDOC = ' ÕÊÌ·' , DOC_TRANS = " & addstring(cDoc) & " WHERE DOC_NO = " & MyParn(xDoc_No.Text)
        
        con.Execute " DELETE FROM NEW1_60H WHERE DOC_NO = " & addstring(cDoc)
        For nRow = 2 To nCountBranch
            sStore = RetZero(nRow, 2)
            con.Execute "  INSERT INTO NEW1_60H (DOC_NO , BRANCH) VALUES        ( " & addstring(cDoc) & " , " & addstring(sStore) & " )  "
        Next nRow
        
        Inform " „  ÕÊÌ· ≈·Ï „” ‰œ  ÕÊÌ· —ﬁ„ " & cDoc
    
        openCardTable
        myUndo
    
    End If
End Sub
Private Sub CMD_RETINV_Click()
    Dim cDoc As String
    Dim aInsert As Variant
    If MsgBox("≈·Ï „” ‰œ „—œÊœ „‘ —Ì«  ", vbYesNo + vbDefaultButton2) = vbYes Then
        aInsert = AddFlag(Empty, "[DATE]", DateSq(Date))
        aInsert = AddFlag(aInsert, "[mosm]", addstring(cPMosm))
        aInsert = AddFlag(aInsert, "[STORE]", addstring(xStore1.BoundText))
        aInsert = AddFlag(aInsert, "[code]", addstring(grid1.TextMatrix(1, 3)))
        aInsert = AddFlag(aInsert, "[isdate]", "getdate()")
        aInsert = AddFlag(aInsert, "[USERSEND]", addstring(cusername))
        aInsert = AddFlag(aInsert, "[Notes_Doc]", addstring(xNotes.Text))
        aInsert = AddFlag(aInsert, "[TRANS_DOC]", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "[ISNEW1]", 1)
        If nFlagPublic = 0 Then
            cDoc = RetZero(Newflag("FILE7_10H", "DOC_NO"))
            aInsert = AddFlag(aInsert, "DOC_NO", addstring(cDoc))
            con.Execute addInsert(aInsert, "FILE7_10h")
            
            con.Execute "  INSERT INTO FILE7_10 (DOC_NO , quant, ITEM, price )  SELECT  " & addstring(cDoc) & " , quant, file1_61.ITEM, file1_10.COST  From FILE1_61 inner join file1_10 on file1_61.item = file1_10.item WHERE file1_61.item <> 0 and DOC_NO = " & MyParn(xDoc_No.Text)
            con.Execute "  UPDATE FILE1_61H SET  ISNEW = 1 , TYPEDOC = '„—œÊœ „‘ —Ì« ' , DOC_TRANS = " & addstring(cDoc) & " WHERE DOC_NO = " & MyParn(xDoc_No.Text)
            
            con.Execute " DELETE FROM NEW7_10H WHERE DOC_NO = " & addstring(cDoc)
            For nRow = 2 To nCountBranch
                sStore = RetZero(nRow, 2)
                con.Execute "  INSERT INTO NEW7_10H (DOC_NO , BRANCH) VALUES        ( " & addstring(cDoc) & " , " & addstring(sStore) & " )  "
            Next nRow
        Else
            cDoc = Newflag_PurchBr("FR7_10H", "DOC_NO", xStore1.BoundText)
            aInsert = AddFlag(aInsert, "DOC_NO", addstring(cDoc))
            con.Execute addInsert(aInsert, "FR7_10h")
            
            con.Execute "  INSERT INTO FR7_10 (DOC_NO , quant, ITEM, price )  SELECT  " & addstring(cDoc) & " , quant, FR1_61.ITEM, file1_10.COST  From FR1_61 inner join file1_10 on FR1_61.item = file1_10.item WHERE FR1_61.item <> 0 and DOC_NO = " & MyParn(xDoc_No.Text)
            con.Execute "  UPDATE FR1_61H SET TYPEDOC = '„—œÊœ „‘ —Ì« ' , DOC_TRANS = " & addstring(cDoc) & " WHERE DOC_NO = " & MyParn(xDoc_No.Text)
        End If
        Inform " „  ÕÊÌ· ≈·Ï „” ‰œ „—œÊœ „‘ —Ì«  —ﬁ„ " & cDoc
        openCardTable
        myUndo
    End If
End Sub
Sub LookTRansNot()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(4, 1)
Set Generalarray(0) = Me

If nFlagPublic = 0 Then
    Generalarray(1) = "SELECT DOC_NO,DATE, CONVERT(VARCHAR(10),[DATE],111),FILE0_40.DESCA,FILE0_40_1.DESCA " & _
                      " FROM (FILE1_61H INNER JOIN FILE0_40 ON FILE1_61H.Store1 = FILE0_40.CODE) INNER JOIN FILE0_40 AS FILE0_40_1 ON FILE1_61H.STORE2 = FILE0_40_1.CODE "
Else
    Generalarray(1) = "SELECT DOC_NO,DATE, CONVERT(VARCHAR(10),[DATE],111),BRANCH_FR.DESCA,' ' " & _
                      " FROM (FR1_61H INNER JOIN BRANCH_FR ON FR1_61H.Store1 = BRANCH_FR.CODE) "
End If
If cBranch = "00" Then
    Generalarray(1) = Generalarray(1) & " WHERE DOC_TRANS IS NULL "
Else
    Generalarray(1) = Generalarray(1) & " WHERE SENT = 0 "
End If
Generalarray(2) = "Order by Date , DOC_NO "
Generalarray(3) = 4200
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ"
listarray(0, 1) = "(@@Doc_No@@6 OR " & _
                  " ##[DATE]##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "„‰ „Œ“‰"
GrdArray(3, 1) = 2000

GrdArray(4, 0) = "≈·Ì „Œ“‰"
GrdArray(4, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
Load oSearchDoc
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1

End Sub

Private Sub cmD_TXT_Click()
'    On Error GoTo myerror
    Dim cFileName As String
    Dim nCount As Double, nitem As String
    If GRID2.Rows > 3 Then
        If MsgBox("ÌÊÃœ «’‰«› „”Ã·… ›Ï «·„” ‰œ - ≈÷«›… ·‰›” «·„” ‰œ", vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        End If
    End If
    If MsgBox("≈÷«›… »Ì«‰«  „‰ „·› TXT Ã—œ ", vbYesNo) = vbYes Then
        Common1.InitDir = ""
        Common1.FileName = ""
        Common1.Filter = "TEXT (*.TXT*)|*.TXT*"
        Common1.ShowOpen
        cFileName = Common1.FileName
        Open cFileName For Input As #1    ' Open file.
        If cFileName <> "" Then
            I = 1
            Do While Not EOF(1)   ' Loop until end of file.
                I = I + 1
                Line Input #1, TextLine      ' Read line into variable.
                If Trim(TextLine) = "" Then Exit Sub
                If Mid(TextLine, 1, 1) <> "" Then
                    aLocal = Split(TextLine, Chr(13))
                    nitem = Trim(aLocal(0))
                    If GetDesca("select item from file1_10 where item = " & MyParn(nitem)) <> "" Then
                        grid1.AddItem ""
                        grid1.TextMatrix(grid1.Rows - 1, 0) = nitem
                        Dim ItemTable As New ADODB.Recordset
                        If ItemTable.State = adStateOpen Then ItemTable.Close
                        ItemTable.Open "SELECT * FROM FILE1_10 WHERE ITEM = " & MyParn(nitem), con, adOpenStatic, adLockReadOnly, adCmdText
                        With grid1
                            If Not (ItemTable.EOF And ItemTable.BOF) Then
                                grid1.TextMatrix(grid1.Rows - 1, 1) = ItemTable!DESCA
                                grid1.TextMatrix(grid1.Rows - 1, 2) = ItemTable!UNIT & ""
                                grid1.TextMatrix(grid1.Rows - 1, 4) = Val(Trim(aLocal(1)))
                            End If
                        End With
                    End If
                End If
            Loop
            grid1.AddItem ""
        End If
    End If
Close #1   ' Close file.
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear

End Sub

