VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form StockAllDoc 
   Caption         =   " ›’Ì·Ï Ã—œ ·› —… «’‰«›"
   ClientHeight    =   9180
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   13260
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
   ScaleHeight     =   9180
   ScaleWidth      =   13260
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   1410
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   495
      Width           =   4065
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2625
         Picture         =   "StockAllDoc.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   45
         Picture         =   "StockAllDoc.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   135
         Width           =   1185
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   45
         TabIndex        =   24
         Top             =   630
         Width           =   3900
         _ExtentX        =   6879
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
         Picture         =   "StockAllDoc.frx":495E
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   240
         Left            =   45
         TabIndex        =   25
         Top             =   1125
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   423
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   6705
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   10650
      Begin VB.TextBox xitem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3015
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1440
         Width           =   1725
      End
      Begin VB.TextBox xmodelfact 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   7425
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1440
         Width           =   1725
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   360
         Left            =   6000
         TabIndex        =   3
         Top             =   1035
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   360
         Left            =   6000
         TabIndex        =   4
         Top             =   630
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   360
         Left            =   1305
         TabIndex        =   5
         Top             =   180
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   360
         Left            =   1305
         TabIndex        =   6
         Top             =   630
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   360
         Left            =   1305
         TabIndex        =   7
         Top             =   1035
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker xdate1 
         Height          =   375
         Left            =   7515
         TabIndex        =   8
         Top             =   180
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   443547649
         CurrentDate     =   43098.0832986111
      End
      Begin MSComCtl2.DTPicker XDATE2 
         Height          =   375
         Left            =   5760
         TabIndex        =   9
         Top             =   180
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   443547649
         CurrentDate     =   43098.0832986111
      End
      Begin VB.Label Label2 
         Caption         =   "—ﬁ„ »«—ﬂÊœ :"
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
         Index           =   2
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1530
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "—ﬁ„ „ÊœÌ· : "
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
         Index           =   0
         Left            =   9135
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1485
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬁ”„ :"
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
         Index           =   1
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   615
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„’‰⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   1755
         TabIndex        =   15
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "„Ã„Ê⁄… „ÊœÌ·«  :"
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
         Index           =   5
         Left            =   9180
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Label Label3 
         Caption         =   "«·„’‰⁄ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9270
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   705
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ê”„ :"
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
         Index           =   6
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·› —… : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label2 
         Caption         =   "«·›—⁄ : "
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
         Index           =   7
         Left            =   4725
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   990
         Width           =   1365
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   8850
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   17639
            MinWidth        =   17639
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   2520
      Top             =   405
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
      Bindings        =   "StockAllDoc.frx":7159
      Height          =   6780
      Left            =   135
      TabIndex        =   1
      Top             =   1935
      Width           =   17205
      _cx             =   30348
      _cy             =   11959
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      BackColorSel    =   12648447
      ForeColorSel    =   4194368
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
      SelectionMode   =   1
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   60
      Top             =   120
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   300
      Top             =   120
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   60
      Top             =   120
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
Attribute VB_Name = "StockAllDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim cString As String
Dim cStr1 As String, cStr2 As String
Private Sub cmdExel_Click()
    ToFileExel grid1
End Sub
Private Sub cmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "≈Ã„«·Ï  ”ÊÌ«  «·Ã—œ " & xStore.text
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.Value, "DD-MM-YYYY") & " ≈·Ï  «—ÌŒ " & Format(xDate2.Value, "DD-MM-YYYY")
    Load PrintGrd
    PrintGrd.DOPRINT Me.grid1, 1, -2, cHead1, cHead2, , False, , 9
    PrintGrd.Show 1
End Sub

Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub

Private Sub CmdExit_Click()
    Unload Me
    Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
    xDate1.Value = "1-1-" & Year(Date)
    xDate2.Value = Format(Date, "dd-mm-yyyy")

    openCon con
    
    Set data1.Recordset = myRecordSet("Select Code,DescA From File0_40 order by Desca", con)
    Set xStore.RowSource = data1
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
    
    Set grid1.DataSource = DATA4
    FIXGRID
    grid1.Rows = 1
End Sub
Private Sub myload()
Dim cWhere As String, cString As String
With grid1
'                           0               1               2                   3               4           5               6               7                   8               9                       10              11                  12
    cString = "SELECT FILE0_10H.DOC_NO, FILE0_10H.DATE, FILE0_40.DESCA, FILE0_100.ITEM , FILE1_10.MOSM , FACT.DESCA  ,FILE1_10.modelfact0 , FILE1_10.DESCA , FILE1_10.COLOR , FILE1_10.SCAL , FILE1_10.COSTITEM ,FILE0_100.REALBAL, FILE1_10.COSTITEM * FILE0_100.REALBAL    " & _
                " FROM   FILE0_100 INNER JOIN FILE0_10H ON FILE0_100.doc_no = FILE0_10H.DOC_NO INNER JOIN FILE0_40 ON FILE0_10H.STORE = FILE0_40.CODE  inner join file1_10 on file1_10.item = file0_100.item inner join fact on fact.code = file1_10.fact where FILE0_10H.DOC_NO is not null "
    If xStore.BoundText <> "" Then cString = cString & " and file0_10H.[STORE]  = " & MyParn(xStore.BoundText)
    If IsDate(xDate1.Value) Then cString = cString & " and date >= " & DateSq(xDate1.Value)
    If IsDate(xDate2.Value) Then cString = cString & " and date <= " & DateSq(xDate2.Value)
    If xFact.BoundText <> "" Then cString = cString & " and file1_10.[fact]  = " & MyParn(xFact.BoundText)
    If xGroup.BoundText <> "" Then cString = cString & " and file1_10.[group]= " & MyParn(xGroup.BoundText)
    If XSECTION.BoundText <> "" Then cString = cString & " and file1_10.[section]= " & Val(XSECTION.BoundText)
    If xMosm.BoundText <> "" Then cString = cString & " and file1_10.[mosm]= " & MyParn(xMosm.BoundText)
    If xStore.BoundText <> "" Then cString = cString & " and file0_10h.[store]= " & MyParn(xStore.BoundText)
    If xitem.text <> "" Then cString = cString & " and file0_100.[item]= " & Val(xitem.text)
    If xModelFact.text <> "" Then cString = cString & " and file1_10.[modelfact0]= " & MyParn(xModelFact.text)
    
    cString = cString & " order by FILE1_10.MOSM , FACT.DESCA  ,FILE1_10.modelfact0, FILE1_10.COLOR , FILE1_10.SCAL"
    Set DATA4.Recordset = myRecordSet(cString, con)
End With
FIXGRID
End Sub
Sub FIXGRID()
    With grid1
    .Cols = 13
    .RowHeight(0) = 1000
    .WordWrap = True
    
    .TextMatrix(0, 0) = "„” ‰œ"
    .TextMatrix(0, 1) = " «—ÌŒ"
    .TextMatrix(0, 2) = "«·„Œ“‰"
    .TextMatrix(0, 3) = "»«—ﬂÊœ"
    .TextMatrix(0, 4) = "„Ê”„"
    .TextMatrix(0, 5) = "„’‰⁄"
    .TextMatrix(0, 6) = "—ﬁ„ „ÊœÌ·"
    .TextMatrix(0, 7) = "«·’‰›"
    .TextMatrix(0, 8) = "«··Ê‰"
    .TextMatrix(0, 9) = "«·„ﬁ«”"
    .TextMatrix(0, 10) = "«· ﬂ·›…"
    .TextMatrix(0, 11) = "«·⁄œœ"
    .TextMatrix(0, 12) = "«·ﬁÌ„…"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1300
    .ColWidth(2) = 2000
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    
    .ColWidth(5) = 2000
    .ColWidth(6) = 1300
    .ColWidth(7) = 3000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    .ColWidth(10) = 1000
    .ColWidth(11) = 1000
    .ColWidth(12) = 1000
    
    .ColDataType(10) = flexDTDouble
    .ColDataType(11) = flexDTDouble
    .ColDataType(12) = flexDTDouble
    .ColDataType(1) = flexDTDate
    
    .ColFormat(10) = "#0.00"
    .ColFormat(12) = "#0.00"

    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .SubtotalPosition = flexSTBelow
    
    If .Rows > 1 Then
        .Subtotal flexSTSum, -1, 11, "#0", vbRed, vbYellow, True, "  "
        .Subtotal flexSTSum, -1, 12, "#0.00", vbRed, vbYellow, True, "  "
        StatusBar1.Panels(1).text = "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ… : " & grid1.Rows - 2
        .TextMatrix(.Rows - 1, 1) = "«·≈Ã„«·Ï"
    End If
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub grid1_DblClick()
    ModelMove.xitem.text = grid1.TextMatrix(grid1.Row, 3)
    ModelMove.xStore.BoundText = xStore.BoundText
    ModelMove.Show
End Sub
