VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmMosm 
   Caption         =   "ÖÈØ ÇáãæÇÓã"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3810
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   3810
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSCommand CMD_EXIT 
      Height          =   540
      Left            =   0
      TabIndex        =   2
      Top             =   4200
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   953
      _Version        =   196610
      Font3D          =   3
      BackColor       =   13822956
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exit"
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   540
      Left            =   2100
      TabIndex        =   1
      Top             =   4200
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   953
      _Version        =   196610
      Font3D          =   3
      BackColor       =   13822956
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ÍÝÜÜÜÜÜÜÜÜÜÜÜÜÜÜÙ"
   End
   Begin VSFlex7LCtl.VSFlexGrid VsMosm 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3765
      _cx             =   6641
      _cy             =   7223
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      BackColor       =   16777215
      ForeColor       =   12582912
      BackColorFixed  =   12632256
      ForeColorFixed  =   16777215
      BackColorSel    =   16777215
      ForeColorSel    =   255
      BackColorBkg    =   13822956
      BackColorAlternate=   16777215
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
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   3
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmMosm.frx":0000
      ScrollTrack     =   -1  'True
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
      OutlineCol      =   1
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
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   4
   End
End
Attribute VB_Name = "FrmMosm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD_EXIT_Click()
    Unload Me
End Sub
Private Sub Form_Load()
With VsMosm
    .Cols = 3
    .FixedCols = 0
    .FixedRows = 1
    .Rows = 1
    .ColWidth(0) = 700
    .ColWidth(1) = 1200
    .ColWidth(2) = 1500
    .TextMatrix(0, 0) = "ãÛáÞ"
    .TextMatrix(0, 1) = "ãæÓã"
    .TextMatrix(0, 2) = "ÈÏÇíÉ ãæÓã"
    
    .ColDataType(0) = flexDTBoolean
    .ColDataType(2) = flexDTDate
    If TMosm.RecordCount > 0 Then
        TMosm.MoveFirst
        Do While Not TMosm.EOF
            .AddItem ""
'            .RowHeight(.Rows - 1) = 300
            .TextMatrix(.Rows - 1, 0) = TMosm!Close
            .TextMatrix(.Rows - 1, 1) = TMosm.MOSM
            .TextMatrix(.Rows - 1, 2) = Format(TMosm!Date, "DD-MM-YYYY")
            TMosm.MoveNext
            If TMosm.EOF Then Exit Do
        Loop
    End If
    TMosm.MoveFirst
End With
End Sub
Private Sub SSCommand1_Click()
    With VsMosm
    For I = 1 To .Rows - 1
        TMosm.FindFirst " MOSM = " & MyParn(.TextMatrix(I, 1))
        If Not TMosm.NoMatch Then
            TMosm.Edit
            TMosm!Close = .TextMatrix(I, 0)
            TMosm!Date = .TextMatrix(I, 2)
            TMosm.Update
        End If
    Next I
    End With
End Sub
