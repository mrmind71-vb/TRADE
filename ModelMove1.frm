VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form ModelMove 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ—ﬂ… „ÊœÌ·"
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
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
   ScaleHeight     =   9825
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   3060
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1080
      Width           =   3660
      Begin VB.CommandButton cmdGo 
         Height          =   510
         Left            =   2430
         Picture         =   "ModelMove1.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Enabled         =   0   'False
         Height          =   510
         Left            =   1230
         Picture         =   "ModelMove1.frx":3059
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   45
         Picture         =   "ModelMove1.frx":5483
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1050
      Left            =   3060
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   3660
      Begin VB.TextBox xdate 
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
         Left            =   90
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   2535
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   90
         TabIndex        =   7
         Top             =   225
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·„Œ“‰ :"
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
         Height          =   195
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ :"
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
         Height          =   195
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   795
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   540
      Top             =   -135
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
   Begin VB.Frame Frame1 
      Height          =   1680
      Left            =   6750
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   45
      Width           =   8385
      Begin VB.TextBox xItem 
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
         Left            =   4815
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1260
         Width           =   2310
      End
      Begin VB.TextBox xModel 
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
         Left            =   4815
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   2310
      End
      Begin MSDataListLib.DataCombo xscal 
         Height          =   315
         Left            =   5760
         TabIndex        =   10
         Top             =   540
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xColor 
         Height          =   315
         Left            =   5760
         TabIndex        =   12
         Top             =   900
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label xdesca 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   180
         Width           =   4695
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·»«—ﬂÊœ :"
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
         Height          =   195
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1305
         Width           =   1035
      End
      Begin VB.Label xSupp 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   900
         Width           =   2985
      End
      Begin VB.Label xFact 
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
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   900
         Width           =   2580
      End
      Begin VB.Label xMosm 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   540
         Width           =   2985
      End
      Begin VB.Label xModelFact 
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
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   540
         Width           =   2580
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "·Ê‰ :"
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
         Height          =   195
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   945
         Width           =   405
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "„ﬁ«” :"
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
         Height          =   195
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   585
         Width           =   585
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   9360
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   675
         Width           =   75
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1035
      End
   End
   Begin VB.TextBox LastOne 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   -555
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1920
      Width           =   405
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   1260
      Top             =   630
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   2385
      Top             =   2565
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
      Height          =   7455
      Left            =   90
      TabIndex        =   24
      Top             =   1800
      Width           =   15090
      _cx             =   26617
      _cy             =   13150
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
      Rows            =   1
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   1
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   25
      Top             =   9390
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12:49 „"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "ModelMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim osearchitem As New Search3
Dim con As New ADODB.Connection
Sub Fillgrd()
Dim loctable As New ADODB.Recordset
cString = "select file1_11.date ,FILE1_12.DESCA,file1_11.doc_id , file1_11.store , file1_11.[type] , sum(file1_11.[in]) as tin , sum(file1_11.out) as tout  ,file1_12.desca AS TYPEDESCA  ,file4_10.desca as supDesca,file0_40.desca as storeDesca , MIN(FILE1_11.PRICE) as lprice, FILE1_12.[ORDER] " & _
          " FROM  (((FILE1_11 INNER JOIN  FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM)  INNER JOIN FILE1_12 ON FILE1_11.TYPE = FILE1_12.CODE)INNER JOIN FILE0_40 ON FILE1_11.STORE = FILE0_40.CODE) LEFT OUTER JOIN FILE4_10 ON FILE1_11.codesup = FILE4_10.CODE "

If IsNumeric(xItem.Text) Then
    cString = cString & turn(cString) & " file1_10.ITEM = " & xItem.Text
End If

If xModel.Text <> "" Then
    cString = cString & turn(cString) & " file1_10.MODEL = " & MyParn(xModel.Text)
End If
If IsDate(xdate.Text) Then
    cString = cString & turn(cString) & " file1_11.date >= " & DateSq(xdate.Text)
End If

If Trim(xStore.BoundText) <> "" Then
    cString = cString & turn(cString) & " file1_11.store = " & MyParn(xStore.BoundText)
End If

If Trim(xscal.BoundText) <> "" Then
    cString = cString & turn(cString) & " file1_10.C_SCAL = " & MyParn(xscal.BoundText)
End If

If Trim(xColor.BoundText) <> "" Then
    cString = cString & turn(cString) & " file1_10.C_COLOR = " & MyParn(xColor.BoundText)
End If


cString = cString & " group by file1_11.date , file1_11.doc_id , file1_11.store , file1_11.[type]  ,file1_12.desca ,file4_10.desca ,file0_40.desca  , FILE1_12.[ORDER] "
cString = cString & cwhere & " ORDER BY FILE1_11.DATE, FILE1_12.[ORDER] "

With grid1
    .Rows = 1
    If IsDate(xdate.Text) Then
       cString2 = "Select sum([IN] - OUT ) as Balance from file1_11 where file1_11.item = " & MyParn(xItem.Text) & _
                  " and file1_11.date < " & DateSq(xdate.Text) & cwhere
       nPrevious = Val(GetDesca(cString2))
       If nPrevious <> 0 Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = "—’Ìœ ﬁ»· " & xdate.Text
            .TextMatrix(.Rows - 1, 3) = nPrevious
       End If
    End If
    loctable.Open cString, con, adOpenStatic, adLockReadOnly, adcdmtext

    Do Until loctable.EOF
         grid1.AddItem ""
         nPrevious = nPrevious + Val(loctable!tIn & "") - Val(loctable!tout & "")
        .TextMatrix(.Rows - 1, 0) = loctable!Desca
        .TextMatrix(.Rows - 1, 1) = Format(Val(loctable!tout & ""), "#0")
        .TextMatrix(.Rows - 1, 2) = Format(Val(loctable!tIn & ""), "#0")
        .TextMatrix(.Rows - 1, 3) = Format(nPrevious, "#0")
        .TextMatrix(.Rows - 1, 4) = Format(loctable!Date, "yyyy/mm/dd")
        .TextMatrix(.Rows - 1, 5) = Format(Val(loctable!lPrice & ""), "fixed")
        .TextMatrix(.Rows - 1, 6) = loctable!DOC_ID & ""
        .TextMatrix(.Rows - 1, 7) = loctable!supDesca & ""
        .TextMatrix(.Rows - 1, 8) = loctable!Type & ""
        loctable.MoveNext
    Loop
End With
End Sub
Sub myProc()
    ActiveControl.Text = osearchitem.grid1.TextMatrix(osearchitem.grid1.Row, 0)
    osearchitem.Hide
End Sub
Function myValid() As Boolean
If Trim(xModel.Text) = "" And Trim(xItem.Text) = "" Then
    MsgBox "€Ì— „”Ã· „ÊœÌ· «Ê »«—ﬂÊœ"
    Exit Function
End If
If (Not IsDate(xdate.Text)) And Trim(xdate.Text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
myValid = True
End Function

Private Sub CMD_ALL_Click()
    If xbarcode.Text = "" Then Exit Sub
    If MsgBox("Õ–›  Õ—ﬂ… «·»«—ﬂÊœ", vbOKCancel + vbDefaultButton2) = vbOK Then
        If MsgBox("Õ–›  „‘ —Ì«  «·»«—ﬂÊœ", vbOKCancel + vbDefaultButton2) = vbOK Then
            con.Execute " DELETE FROM FILE7_20 WHERE ITEM = " & Val(xbarcode.Text)
        End If
    
        If MsgBox("Õ–›  „—œÊœ „‘ —Ì«  «·»«—ﬂÊœ", vbOKCancel + vbDefaultButton2) = vbOK Then
            con.Execute " DELETE FROM FILE7_30 WHERE ITEM = " & Val(xbarcode.Text)
        End If
    
        If MsgBox("Õ–›  Ã—œ «·»«—ﬂÊœ", vbOKCancel + vbDefaultButton2) = vbOK Then
            con.Execute " DELETE FROM FILE0_10 WHERE ITEM = " & Val(xbarcode.Text)
        End If
    
        If MsgBox("Õ–›  »Ì«‰«  «·»«—ﬂÊœ", vbOKCancel + vbDefaultButton2) = vbOK Then
            con.Execute " DELETE FROM FILE1_10 WHERE ITEM = " & Val(xbarcode.Text)
        End If
    End If
    CmdGo_Click
End Sub

Private Sub CmdGo_Click()
If Not myValid Then Exit Sub
Fillgrd
'xBal.Caption = Format(Val(grid1.TextMatrix(grid1.Rows - 1, 3)), "#0.00")
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub CMD_PRINT_Click()
'    Load rpitem2
'    rpitem2.xItem.Text = xItem.Text
'    rpitem2.xdate1.Text = xdate.Text
'    rpitem2.xStore.BoundText = xStore.BoundText
'    rpitem2.Show 1
    Dim cHead1 As String
    Dim cHead2 As String
    Dim cHead3 As String

    If xSupp.Caption = "" Then
        cHead1 = ArbString(" ›’Ì·Ï Õ—ﬂ… „ÊœÌ· —ﬁ„ " & xModelFact.Caption & " „’‰⁄ " & xFact.Caption & " „Ê”„ " & xMosm.Caption)
    Else
        cHead1 = ArbString(" ›’Ì·Ï Õ—ﬂ… „ÊœÌ· —ﬁ„ " & xModelFact.Caption & " „’‰⁄ " & xFact.Caption & " „Ê”„ " & xMosm.Caption & " „ﬂ » " & xSupp.Caption)
    End If
    If xscal.BoundText <> "" Then cHead2 = cHead2 & ArbString(" „ﬁ«” " & xscal.Text)
    If xColor.BoundText <> "" Then cHead2 = cHead2 & ArbString(" ·Ê‰ " & xColor.Text)
    If xStore.BoundText <> "" Then cHead3 = cHead3 & ArbString(" „Œ“‰  " & xStore.Text)
    If IsDate(xdate.Text) Then cHead3 = cHead3 & ArbString(" „‰  «—ÌŒ " & xdate.Text)
    
    Load PrintGrd
    PrintGrd.doprint grid1, 0.8, 1, cHead1, cHead2, cHead3, False, , 9
    PrintGrd.Show 1

End Sub
Private Sub Form_Load()
openCon con

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FILE0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

xStore.BoundText = sBranchCode

If sBranchCode <> "00" Then xStore.Enabled = False
'DATA3.ConnectionString = strCon
'Set xscal.RowSource = DATA3
'
'DATA3.ConnectionString = strCon
'Set xColor.RowSource = DATA3

With grid1
.FixedCols = 0

grid1.Cols = 9
.TextMatrix(0, 0) = "»Ì«‰"
.TextMatrix(0, 1) = "’«œ—"
.TextMatrix(0, 2) = "Ê«—œ"
.TextMatrix(0, 3) = "—’Ìœ"
.TextMatrix(0, 4) = " «—ÌŒ"
.TextMatrix(0, 5) = "”⁄—"
.TextMatrix(0, 6) = "„” ‰œ"
.TextMatrix(0, 7) = "≈”„"


grid1.ColWidth(0) = 3000
grid1.ColWidth(1) = 1000
grid1.ColWidth(2) = 1000
grid1.ColWidth(3) = 1000
grid1.ColWidth(4) = 1500
grid1.ColWidth(5) = 1000
grid1.ColWidth(6) = 1700
grid1.ColWidth(7) = 3500
grid1.ColHidden(8) = True
End With
For i = 0 To grid1.Cols - 1
    grid1.ColAlignment(i) = flexAlignRightCenter
Next


DATA2.ConnectionString = strCon
DATA2.RecordSource = "SELECT  C_SCAL , SCAL FROM FILE1_10 WHERE 1 = 2"
Set xscal.RowSource = DATA2
xscal.ListField = "SCAL"
xscal.BoundColumn = "C_SCAL"
DATA2.Refresh

data3.ConnectionString = strCon
data3.RecordSource = "SELECT C_COLOR, COLOR FROM FILE1_10 WHERE 1 = 2"
xColor.ListField = "COLOR"
xColor.BoundColumn = "C_COLOR"
Set xColor.RowSource = data3
data3.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Unload osearchitem
Err.Clear
End Sub

Private Sub grid1_dblClick()
    'bEdit = RetSec("xPurchase")
    'If bEdit Then
    Select Case grid1.TextMatrix(grid1.Row, 8)
    Case "2", "7"
'        Load Purchasefrm
'        Purchasefrm.myPublic = IIf(grid1.TextMatrix(grid1.Row, 8) = "2", 0, 1)
'        Purchasefrm.myproc2 grid1.TextMatrix(grid1.Row, 6)
'        Purchasefrm.Show
    Case "3", "6"
'        SALESFRM.myPublic = IIf(grid1.TextMatrix(grid1.Row, 8) = "6", 0, 1)
'        SALESFRM.sDoc_No = grid1.TextMatrix(grid1.Row, 6)
'        SALESFRM.Show
    End Select
    'Else
    '    MsgBox "·Ì” ·œÌﬂ «·’·«ÕÌ… ·œŒÊ· „·› «·„‘ —Ì« "
    '    Exit Sub
    'End If
End Sub
Private Sub CalcScal(sModel)
Dim cString As String, sValue As String
cValue = xscal.BoundText
If Trim(sModel) = "" Then
    cString = "SELECT  C_SCAL , SCAL FROM FILE1_10 WHERE 1 = 2"
Else
    cString = "SELECT  C_SCAL , SCAL FROM FILE1_10"
    cString = cString & turn(cString) & "MODEL = " & MyParn(sModel)
    cString = cString & " GROUP BY C_SCAL , SCAL ORDER BY C_SCAL "
End If
If DATA2.RecordSource <> cString Then
    DATA2.RecordSource = cString
    DATA2.Refresh
End If
xscal.BoundText = sValue
If Not xscal.MatchedWithList Then xscal.BoundText = ""
End Sub
Private Sub CalcColor(sModel)
Dim cString As String, sValue As String
cValue = xColor.BoundText
If Trim(sModel) = "" Then
    cString = "SELECT  C_COLOR , COLOR FROM FILE1_10 WHERE 1 = 2"
Else
    cString = "SELECT  C_COLOR , COLOR FROM FILE1_10"
    cString = cString & turn(cString) & "MODEL = " & MyParn(sModel)
    cString = cString & " GROUP BY C_COLOR , COLOR ORDER BY C_COLOR "
End If
If data3.RecordSource <> cString Then
    data3.RecordSource = cString
    data3.Refresh
End If
xColor.BoundText = sValue
If Not xColor.MatchedWithList Then xColor.BoundText = ""
End Sub
Private Sub xItem_Validate(Cancel As Boolean)
If IsNumeric(xItem.Text) Then
    If Trim(xModel.Text) <> "" Then
        xModel.Text = ""
        xModel_LostFocus
    End If
    xDesca.Caption = GetDesca("SELECT DESCA FROM FILE1_10 WHERE ITEM =" & MyParn(xItem.Text))
Else
    If Trim(xModel.Text) = "" Then xDesca.Caption = ""
End If
End Sub

Private Sub xModel_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ModelLookupAll Me, osearchitem
End Sub
Private Sub xModel_LostFocus()
If Trim(xModel.Text) = "" Then
    xDesca.Caption = ""
    xFact.Caption = ""
    xMosm.Caption = ""
    xModelFact.Caption = ""
    xSupp.Caption = ""
Else
    Dim cString As String, aRet As Variant
    cString = "SELECT DESCA,FACT, MOSM , MODELFACT , SUPP FROM FILE1_10H"
    cString = cString & turn(cString) & "MODEL = " & MyParn(xModel.Text)
    aRet = aGetDesca(cString)
    If UBound(aRet) > 0 Then
        xDesca.Caption = aRet(1)
        xFact.Caption = GetDesca("SELECT DESCA FROM FACT WHERE CODE = " & MyParn(aRet(2)))
        xMosm.Caption = aRet(3)
        xModelFact.Caption = DelZero(aRet(4))
        xSupp.Caption = aRet(5) & ""
        CalcColor xModel.Text
        CalcScal xModel.Text
    Else
        CalcColor ""
        CalcScal ""
    End If
End If
End Sub
Private Sub xStore_Click(Area As Integer)
If Not cmdGo.Enabled Then cmdGo.Enabled = True
End Sub
Private Sub xStore_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdGo_Click
End Sub
Private Sub xstore_LostFocus()
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
