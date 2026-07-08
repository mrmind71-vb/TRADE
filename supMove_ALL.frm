VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form supMove_ALL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ—ﬂ… «·„Ê—œÌ‰"
   ClientHeight    =   9675
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
   ScaleHeight     =   9675
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Height          =   1545
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   765
      Width           =   2490
      Begin VB.CommandButton NEW_MOSM 
         BackColor       =   &H00E0E0E0&
         Caption         =   " ”ÃÌ· „Ê”„ ÃœÌœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   180
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton CMD_MOSM 
         BackColor       =   &H00E0E0E0&
         Caption         =   "—’Ìœ √Ê· ··„Ê”„ ··„Ê—œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   855
         Width           =   2325
      End
   End
   Begin VB.Frame Frame5 
      Height          =   2355
      Left            =   8460
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1080
      Width           =   6675
      Begin VSFlex7Ctl.VSFlexGrid VsTot 
         Height          =   2040
         Left            =   90
         TabIndex        =   19
         Top             =   180
         Width           =   6495
         _cx             =   11456
         _cy             =   3598
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   12648384
         ForeColorSel    =   64
         BackColorBkg    =   12648447
         BackColorAlternate=   14737632
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
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
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
         AutoResize      =   -1  'True
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
         Editable        =   2
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
   End
   Begin VB.Frame Frame4 
      Height          =   2355
      Left            =   2700
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   765
      Visible         =   0   'False
      Width           =   5730
      Begin VSFlex7Ctl.VSFlexGrid VsMosm 
         Height          =   2040
         Left            =   90
         TabIndex        =   20
         Top             =   180
         Width           =   5550
         _cx             =   9790
         _cy             =   3598
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   6768636
         ForeColorSel    =   64
         BackColorBkg    =   12648447
         BackColorAlternate=   14737632
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
         Rows            =   2
         Cols            =   3
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
         AutoResize      =   -1  'True
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
         Editable        =   2
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
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   6840
      Top             =   900
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
      Height          =   960
      Left            =   8460
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   90
      Width           =   6675
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
         Height          =   330
         Left            =   270
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox xdate1 
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
         Left            =   3600
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox XCODE 
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
         TabIndex        =   0
         Top             =   180
         Width           =   915
      End
      Begin VB.Label xMosm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Õ Ì  «—ÌŒ"
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
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ"
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
         Left            =   5445
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   630
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "ﬂÊœ «·„Ê—œ"
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
         Left            =   5775
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   825
      End
      Begin VB.Label xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   180
         Width           =   3435
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
      TabIndex        =   5
      Top             =   1920
      Width           =   405
   End
   Begin VSFlex7Ctl.VSFlexGrid invGrid 
      Height          =   5730
      Left            =   135
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3465
      Width           =   15000
      _cx             =   26458
      _cy             =   10107
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic"
         Size            =   11.25
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
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   90
      Width           =   5730
      Begin VB.CommandButton CmdGo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "≈ŸÂ«— «·Õ—ﬂ…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   180
         Width           =   1920
      End
      Begin VB.CommandButton CMD_PRINT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ÿ»«⁄… "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   180
         Width           =   1785
      End
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Œ—ÊÃ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   180
         Width           =   1680
      End
   End
   Begin VB.Frame Frame3 
      Height          =   645
      Left            =   5940
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   90
      Width           =   2490
      Begin VB.Label xBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         TabIndex        =   12
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·—’Ìœ"
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
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   525
      End
   End
End
Attribute VB_Name = "supMove_ALL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim ClientTable As New ADODB.Recordset
Sub Fillgrd()
Dim loctable As New ADODB.Recordset
If XCODE.Text = "" Then
    MsgBox "„—«Ã⁄… ﬂÊœ «·„Ê—œ "
    Exit Sub
End If
cString = "select FILE4_11_2.*,FILE4_12.desca as MoveDesca  " & _
          " From FILE4_11_2 left join FILE4_12 on FILE4_11_2.[type] = FILE4_12.code WHERE "
cString = cString & turn(cString) & " FILE4_11_2.code = " & MyParn(XCODE.Text)
'cString = cString & turn(cString) & " FILE4_11_2.MOSM = " & MyParn(xMosm.Caption)

'If IsDate(xdate1.Text) Then
'    cString = cString & turnFound(cString) & " FILE4_11_2.date >= " & DateSq(xdate1.Text)
'End If
'
'If IsDate(XDATE2.Text) Then
'    cString = cString & turnFound(cString) & " FILE4_11_2.date <= " & DateSq(XDATE2.Text)
'End If

cString = cString & " order by MYINDEX , Date,FILE4_12.[ORDER],FILE4_11_2.SAL,FILE4_11_2.doc_id"

With invGrid
    .Rows = 1
'    If IsDate(xdate1.Text) Then
'       cString2 = "Select sum([SAL] - PAY) as Balance from FILE4_11_2 where FILE4_11_2.CODE = " & MyParn(XCODE.Text) & _
'                  " and FILE4_11_2.date < " & DateSq(xdate1.Text)
'       nPrevious = Val(GetDesca(cString2))
'       If nPrevious <> 0 Then
'            .AddItem ""
'            .TextMatrix(.Rows - 1, 0) = "—’Ìœ ﬁ»· " & xdate1.Text
'            .TextMatrix(.Rows - 1, 4) = nPrevious
'       End If
'    End If
    loctable.Open cString, con, adOpenStatic, adLockReadOnly, adcdmtext
'    Inform "7"
    Do Until loctable.EOF
         invGrid.AddItem ""
         nPrevious = nPrevious + Val(loctable!sal & "") - Val(loctable!PAY & "")
        .TextMatrix(.Rows - 1, 0) = loctable!DESCA & ""
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Date, "yyyy/mm/dd")
        .TextMatrix(.Rows - 1, 2) = loctable!doc_ID & ""
        .TextMatrix(.Rows - 1, 3) = Format(TurnValue(Val(loctable!PAY & ""), 0, ""), "#0.00")
        .TextMatrix(.Rows - 1, 4) = Format(TurnValue(Val(loctable!sal & ""), 0, ""), "#0.00")
        .TextMatrix(.Rows - 1, 5) = Format(nPrevious, "#0.00")
        .TextMatrix(.Rows - 1, 6) = loctable!Type & ""
        
        .TextMatrix(.Rows - 1, 7) = loctable!INV_NO & ""
        .TextMatrix(.Rows - 1, 8) = loctable!STORE & ""
        loctable.MoveNext
    Loop
    If .Rows > 1 Then
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, -1, 3, "#0.00", vbYellow, vbRed, True, "  "
        .Subtotal flexSTSum, -1, 4, "#0.00", vbYellow, vbRed, True, "  "
        .TextMatrix(.Rows - 1, 0) = "«·«Ã„«·Ì"
        .TextMatrix(.Rows - 1, 5) = Format(Round(nPrevious, 2), "#0.00")
    End If
End With
xBal.Caption = Format(nPrevious, "#0.00")
invGrid.ColComboList(8) = StrListA("SELECT * from BRANCH ")

'Inform "9"
FillTot
'Inform "10"
End Sub
Sub myProc()
ActiveControl.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
Search3.Hide
End Sub
Function MYVALID() As Boolean
'If xMosm.Caption = "" Then
'
'    MsgBox "«·„Ê”„ €Ì— „Õœœ "
'    Exit Function
'End If

If XCODE.Text = "" Then
    MsgBox "ﬂÊœ «·„Ê—œ €Ì— „”Ã·"
    Exit Function
End If
ClientTable.Find " code = " & MyParn(XCODE.Text), , adSearchForward, adBookmarkFirst
If ClientTable.EOF Then
    MsgBox "ﬂÊœ «·„Ê—œ €Ì— ’ÕÌÕ"
    Exit Function
End If
'If (Not IsDate(xdate1.Text)) And Trim(xdate1.Text) <> "" Then
'    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
'    Exit Function
'End If
'If (Not IsDate(XDATE2.Text)) And Trim(XDATE2.Text) <> "" Then
'    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
'    Exit Function
'End If
MYVALID = True
End Function
Private Sub cmdcorect_Click()

End Sub
Private Sub CMD_MOSM_Click()
    If XCODE.Text = "" Then Exit Sub
    FrmFMosm.Show 1
    cStr1 = "SELECT  mosm, DATE, f_bal, BAL FROM   BAL_MOSM WHERE CODE = " & MyParn(XCODE.Text) & " ORDER BY DATE DESC "
    data1.RecordSource = cStr1
    data1.Refresh
    FixGridMosm
End Sub
Private Sub CMD_PRINT_Click()
Dim nBalance As Double, nRow As Integer
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim cMosmDesca As String, cPhone   As String
cPhone = GetDesca("SELECT PHONE1 FROM FILE4_10 WHERE CODE = " & MyParn(XCODE.Text))
cPhone = cPhone & " " & GetDesca("SELECT FAX FROM FILE4_10 WHERE CODE = " & MyParn(XCODE.Text))
Dim nBal As Double
Dim n10 As Double, n11 As Double, n12 As Double, n13 As Double, n14 As Double
Dim n15 As Double, n16 As Double, n17 As Double, n18 As Double, n19 As Double, n20 As Double
Dim cWhere As String
cWhere = " CODE = " & MyParn(XCODE.Text) & " And MOSM = " & MyParn(xMosm.Caption)


n2 = Val(GetDesca("SELECT SUM(TAX ) FROM FILE7_20H WHERE " & cWhere) & "")
n3 = Val(GetDesca("SELECT SUM(DISCOUNT) FROM FILE7_20H WHERE " & cWhere) & "")
n4 = Val(GetDesca("SELECT SUM(SAL)      FROM FILE4_11_2  WHERE " & cWhere & " AND [TYPE] = '4' ") & "")
n1 = n4 - n2 + n3

n6 = Val(GetDesca("SELECT SUM(TAX ) FROM FILE7_10H WHERE " & cWhere) & "")
n7 = Val(GetDesca("SELECT SUM(DISCOUNT) FROM FILE7_10H WHERE " & cWhere) & "")
n8 = Val(GetDesca("SELECT SUM(pay)      FROM FILE4_11_2  WHERE " & cWhere & " AND [TYPE] = '5' ") & "")
n5 = n8 - n6 + n7
 
n9 = Val(GetDesca("SELECT SUM(PAY -SAL )       FROM FILE4_11_2  WHERE " & cWhere & " AND [TYPE] > '5' ") & "")
n10 = Val(GetDesca("SELECT SUM(SAL -PAY )       FROM FILE4_11_2  WHERE " & cWhere) & "")
 
cString = "SELECT * FROM FILE4_11_2  " & _
          " WHERE MOSM = " & MyParn(xMosm.Caption) & " AND CODE = " & MyParn(XCODE.Text) & " ORDER BY MYINDEX , DATE , pay , [type] "
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
cMosmDesca = GetDesca("SELECT DESCA FROM MOSM WHERE MOSM = " & MyParn(xMosm.Caption))
With sourcetable
    Do While Not .EOF
        temptable.AddNew
        temptable!Date1 = !Date
        temptable!str1 = !doc_ID
        temptable!str2 = !DESCA
        temptable!val1 = !TITEM
        temptable!val3 = !tax
        temptable!val2 = !disc
        temptable!str3 = TurnValue(!INV_NO, "", Null)
        
        Select Case !Type
            Case "4"
                temptable!val15 = !sal
            Case "5"
                temptable!val16 = !PAY
            Case Else
                temptable!VAL4 = !PAY - !sal
        End Select
        nBal = nBal + Val(!sal & "") - Val(!PAY & "")
        temptable!val5 = nBal
        temptable!str8 = cMosmDesca
        
        If Val(!TITEM) <> 0 Then temptable!VAL7 = Round(Val(!tax & "") / Val(!TITEM & "") * 100, 2)
        
        temptable!STR7 = xDesca.Caption
        temptable!str4 = cPhone
        
        temptable!STR11 = n1
        temptable!str12 = n2
        temptable!str13 = n3
        temptable!str14 = n4
        temptable!str15 = n5
        temptable!str16 = n6
        temptable!str17 = n7
        temptable!str18 = n8
        temptable!str19 = n9
        temptable!STR20 = n10
        
        
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
main.REPORT1.ReportFileName = App.Path & "\Reports\rep_203.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub cmdGo_Click()
'   On Error GoTo myerror
    Dim dFdate As Date
'    If xMosm.Caption = "" And VsMosm.Rows > 1 Then
'        xMosm.Caption = VsMosm.TextMatrix(1, 0)
'    Else
'        If GetDesca("SELECT CODE FROM FILE4_22 WHERE MOSM = " & MyParn(cPMosm) & " AND CODE = " & MyParn(XCODE.Text)) = "" Then
''            Inform "1"
'            dFdate = GetDesca("SELECT DATE FROM MOSM WHERE MOSM = " & MyParn(cPMosm))
''            Inform "2"
'            If IsDate(dFdate) Then
'                con.Execute " INSERT INTO FILE4_22 (MOSM,CODE,DATE,F_BAL) " & _
'                       "Values(" & _
'                       addstring(cPMosm) & "," & _
'                       addstring(XCODE.Text) & "," & _
'                       addDate(dFdate) & "," & _
'                       0 & _
'                       ")"
'                cStr1 = "SELECT  mosm, DATE, f_bal, BAL FROM   BAL_MOSM WHERE CODE = " & MyParn(XCODE.Text) & " ORDER BY DATE DESC "
'                data1.RecordSource = cStr1
'                data1.Refresh
' '               Inform "3"
'
'                FixGridMosm
'            End If
'        End If
'    End If
    If Not MYVALID Then Exit Sub
'    Inform "4"
'
    Fillgrd
'   Inform "6"
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
openCon con

'cStr1 = " INSERT INTO file4_22 (code, mosm, DATE, f_bal) SELECT     FILE4_11_2.code, FILE4_11_2.mosm, GETDATE() AS Expr1, 0 AS Expr2 " & _
'        " FROM         FILE4_10 INNER JOIN FILE4_11_2 ON FILE4_10.CODE = FILE4_11_2.code LEFT OUTER JOIN file4_22 AS file4_22_1 ON FILE4_11_2.mosm = file4_22_1.mosm AND FILE4_11_2.code = file4_22_1.code " & _
'        " WHERE     (file4_22_1.code IS NULL) AND (FILE4_11_2.code IS NOT NULL) and FILE4_11_2.mosm is not  nullGROUP BY FILE4_11_2.code, FILE4_11_2.mosm "
'CON.Execute cStr1

cStr1 = " INSERT INTO file4_22 (code, mosm, DATE, f_bal) SELECT     FILE4_11_2.code, FILE4_11_2.mosm, GETDATE() AS Expr1, 0 AS Expr2 " & _
        " FROM         FILE4_10 INNER JOIN FILE4_11_2 ON FILE4_10.CODE = FILE4_11_2.code LEFT OUTER JOIN file4_22 AS file4_22_1 ON FILE4_11_2.mosm = file4_22_1.mosm AND FILE4_11_2.code = file4_22_1.code " & _
        " WHERE     (file4_22_1.code IS NULL) AND (FILE4_11_2.code IS NOT NULL) and FILE4_11_2.mosm is not  null GROUP BY FILE4_11_2.code, FILE4_11_2.mosm "
con.Execute cStr1

cStr1 = "DELETE FROM file4_22 FROM  FILE4_11_2 RIGHT OUTER JOIN file4_22 ON FILE4_11_2.mosm = file4_22.mosm AND FILE4_11_2.code = file4_22.code Where FILE4_11_2.code Is Null "
con.Execute cStr1

ClientTable.Open "FILE4_10", con, adOpenStatic, adLockReadOnly, adCmdTable

Set VsMosm.DataSource = data1
data1.ConnectionString = strCon


With invGrid
invGrid.Cols = 9
.TextMatrix(0, 0) = "»Ì«‰"
.TextMatrix(0, 1) = " «—ÌŒ"
.TextMatrix(0, 2) = "„” ‰œ"
.TextMatrix(0, 3) = "„œÌ‰"
.TextMatrix(0, 4) = "œ«∆‰"
.TextMatrix(0, 5) = "—’Ìœ"
.TextMatrix(0, 7) = "›« Ê—…"
.TextMatrix(0, 8) = "«·›—⁄ "


invGrid.ColWidth(0) = 4000
invGrid.ColWidth(1) = 1500
invGrid.ColWidth(2) = 1500
invGrid.ColWidth(3) = 1500
invGrid.ColWidth(4) = 1500
invGrid.ColWidth(5) = 1500
invGrid.ColWidth(6) = 0
invGrid.ColWidth(7) = 800
invGrid.ColWidth(8) = 1500
invGrid.ColComboList(8) = StrListA("SELECT * from BRANCH ")
End With
FixGridMosm
FixGridVsTot
VsMosm.Rows = 1
For i = 0 To invGrid.Cols - 1
    invGrid.ColAlignment(i) = flexAlignRightCenter
Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Unload Search3
Err.Clear
End Sub
Private Sub invGrid_dblClick()
    If bopt2 Then
    Select Case invGrid.TextMatrix(invGrid.Row, 6)
        Case "4", "5"
            cDoc_No = invGrid.TextMatrix(invGrid.Row, 2)
            If invGrid.TextMatrix(invGrid.Row, 6) = "4" Then
'               Purchasefrm.bEdit = RetSec(main.tmpurchases.Name)
            Else
'                Purchasefrm.bEdit = RetSec(main.tmRETPURCH.Name)
            End If
            Purchasefrm.myPublic = IIf(invGrid.TextMatrix(invGrid.Row, 6) = "4", 0, 1)
            Purchasefrm.sDoc_no = cDoc_No
            Purchasefrm.Show
    End Select
    End If
End Sub
Private Sub NEW_MOSM_Click()
'    FrmOpenMosm.Show 1
    
'    cStr1 = "SELECT  mosm, DATE, f_bal, BAL FROM   BAL_MOSM WHERE CODE = " & MyParn(xCode.Text) & " ORDER BY DATE DESC "
'    data1.RecordSource = cStr1
'    data1.Refresh
'    FixGridMosm

End Sub
Private Sub VsMosm_DBLClick()
    xMosm.Caption = VsMosm.TextMatrix(VsMosm.Row, 0)
    cmdGo_Click
End Sub
Private Sub xcode_Validate(Cancel As Boolean)
    If XCODE.Text = "" Then Exit Sub
    If GetDesca("SELECT CODE FROM FILE4_10 WHERE code = " & MyParn(XCODE.Text)) = "" Then
        cCode = GetDesca("SELECT CODE FROM FILE4_10 WHERE SUBCODE = " & MyParn(XCODE.Text))
        XCODE.Text = cCode
    End If
End Sub

Private Sub XDATE1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdGo_Click
End Sub
Private Sub xCode_Change()
    invGrid.Rows = 1
    CmdGo.Enabled = Trim(XCODE.Text) <> ""
End Sub
Private Sub xcode_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then CmdGo_Click
End Sub
Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
    If KeyCode = 112 Then CardLookup
End Sub
Private Sub xCode_LostFocus()
    LoadSuplData
End Sub
Private Sub xStore_Click(Area As Integer)
    If Not CmdGo.Enabled Then CmdGo.Enabled = True
End Sub
Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FILE4_10"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1
End Sub
Sub FixGridMosm()
With VsMosm
    .Editable = flexEDNone
    .Cols = 4
    .TextMatrix(0, 0) = "„Ê”„"
    .TextMatrix(0, 1) = " «—ÌŒ √Ê·"
    .TextMatrix(0, 2) = "—’Ìœ √Ê·"
    .TextMatrix(0, 3) = "—’Ìœ √Œ—"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColFormat(1) = "dd-mm-yyyy"
    .ColFormat(2) = "#0.00"
    .ColFormat(3) = "#0.00"

    .ColDataType(1) = flexDTDate
    .ColDataType(2) = flexDTDouble
    .ColDataType(3) = flexDTDouble

End With
End Sub
Sub FixGridVsTot()

With VsTot
    .Editable = flexEDNone
    .Cols = 4
    .Rows = 6
    .FixedCols = 1
    .FixedRows = 1
    .ForeColorFixed = vbBlue
    For i = 0 To .Cols - 1
        .ColWidth(i) = 1500
    Next i
    .ColAlignment(0) = flexAlignCenterCenter
    .Cell(flexcpAlignment, 0, 1) = flexAlignCenterCenter
    .Cell(flexcpAlignment, 0, 2) = flexAlignCenterCenter
    .Cell(flexcpAlignment, 0, 3) = flexAlignCenterCenter
    
    .TextMatrix(0, 1) = "«’‰‹«›"
    .TextMatrix(0, 2) = "Œ’‹‹‹„"
    .TextMatrix(0, 3) = "’«›‹‹Ï"
    
    .TextMatrix(1, 0) = "—’Ìœ √Ê·"
    .TextMatrix(2, 0) = "›Ê« Ì‹‹—"
    .TextMatrix(3, 0) = "„— Ã⁄‹‹« "
    .TextMatrix(4, 0) = "”‹‹‹‹œ«œ"
    .TextMatrix(5, 0) = "—’Ì‹‹‹‹œ"
End With
End Sub
Sub FillTot()
    Dim nV1 As Double, nV2 As Double, nV3 As Double, nV4 As Double, nV5 As Double
    nV1 = Val(GetDesca("SELECT sum(DISCOUNT) FROM FILE7_20H WHERE CODE = " & MyParn(XCODE.Text) & " AND MOSM = " & MyParn(xMosm.Caption)) & "")
    nV2 = Val(GetDesca("SELECT sum(DISCOUNT) FROM FILE7_10H WHERE CODE = " & MyParn(XCODE.Text) & " AND MOSM = " & MyParn(xMosm.Caption)) & "")
    nV3 = Val(GetDesca("SELECT SUM(SAL)  FROM FILE4_11_2 WHERE [TYPE] = '4' AND CODE = " & MyParn(XCODE.Text) & " AND MOSM = " & MyParn(xMosm.Caption)) & "")
    nV4 = Val(GetDesca("SELECT SUM(PAY)  FROM FILE4_11_2 WHERE [TYPE] = '5' AND CODE = " & MyParn(XCODE.Text) & " AND MOSM = " & MyParn(xMosm.Caption)) & "")
    nV5 = Val(GetDesca("SELECT SUM(PAY - SAL )  FROM FILE4_11_2 WHERE ([TYPE] = '7' OR [TYPE] = 'A' OR [TYPE] = 'C' ) AND CODE = " & MyParn(XCODE.Text) & " AND MOSM = " & MyParn(xMosm.Caption)) & "")
    nBal = Val(GetDesca("SELECT SUM(sal - pay )  FROM FILE4_11_2 WHERE CODE = " & MyParn(XCODE.Text) & " AND MOSM = " & MyParn(xMosm.Caption)) & "")
    
    With VsTot
        .TextMatrix(1, 3) = Format(VsMosm.TextMatrix(VsMosm.Row, 2), "#0.00")
        .TextMatrix(2, 2) = Format(nV1, "#0.00")
        .TextMatrix(2, 3) = Format(nV3, "#0.00")
        .TextMatrix(2, 1) = Format(Val(.TextMatrix(2, 3)) + Val(.TextMatrix(2, 2)), "#0.00")
        .TextMatrix(3, 2) = Format(nV2, "#0.00")
        .TextMatrix(3, 3) = Format(nV4, "#0.00")
        .TextMatrix(3, 1) = Format(Val(.TextMatrix(3, 3)) + Val(.TextMatrix(3, 2)), "#0.00")
        .TextMatrix(4, 3) = Format(nV5, "#0.00")
        .TextMatrix(5, 3) = Format(nBal, "#0.00")
    
    End With
End Sub
Sub LoadSuplData()
    xDesca.Caption = ""
    If Trim(XCODE.Text) = "" Then Exit Sub
    XCODE.Text = RetZero(XCODE.Text, 3)
    ClientTable.Find "code = " & MyParn(XCODE.Text), , adSearchForward, adBookmarkFirst
    If Not ClientTable.EOF Then xDesca.Caption = ClientTable!DESCA & ""
    cStr1 = "SELECT  mosm, DATE, f_bal, BAL FROM   BAL_MOSM WHERE CODE = " & MyParn(XCODE.Text) & " ORDER BY DATE DESC "
    data1.RecordSource = cStr1
    data1.Refresh
    FixGridMosm

End Sub

