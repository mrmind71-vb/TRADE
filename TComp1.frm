VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Begin VB.Form TComp1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ﬁ«—Ì— ⁄«„…"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   11100
   ScaleWidth      =   20400
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "› —… À«‰Ì…"
      Height          =   1365
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   75
      Width           =   3015
      Begin VB.TextBox xDate11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   547
         Width           =   1155
      End
      Begin VB.TextBox xDate22 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   930
         Width           =   1155
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„Ê”„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1800
         TabIndex        =   16
         Top             =   150
         Width           =   555
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   270
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   525
         Width           =   675
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   270
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   930
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "› —…  √Ê·Ï"
      Height          =   1365
      Left            =   8775
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   75
      Width           =   3015
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   547
         Width           =   1155
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   930
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„Ê”„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2025
         TabIndex        =   15
         Top             =   150
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   270
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   270
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1005
         Width           =   675
      End
   End
   Begin VB.Frame TCOMP_1 
      Height          =   645
      Left            =   1230
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   75
      Width           =   4305
      Begin VB.CommandButton cmd_print 
         Caption         =   "ÿ»«⁄…"
         Height          =   375
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   180
         Width           =   1005
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   375
         Left            =   2835
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1005
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid invGrid 
      Height          =   6315
      Left            =   225
      TabIndex        =   13
      Top             =   1500
      Width           =   11565
      _cx             =   20399
      _cy             =   11139
      _ConvInfo       =   1
      Appearance      =   1
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
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   128
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
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
      SelectionMode   =   1
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
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
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
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "TComp1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CMD_PRINT_Click()
    Load PrintGrd
    PrintGrd.doprint Me.invGrid, 1, -2, "ÿ»«⁄… „ﬁ«—‰… »Ì‰ " & xDate1.Text & " - " & xDate2.Text, " „‰ " & xDate11.Text & " - " & xDate22.Text, , , , 10
    PrintGrd.Show 1
End Sub
Private Sub CmdApply_Click()
    If Not IsDate(xDate1.Text) Then
        MsgBox "—«Ã⁄ «· «—ÌŒ 1"
        Exit Sub
    End If
    If Not IsDate(xDate2.Text) Then
        MsgBox "—«Ã⁄ «· «—ÌŒ 2"
        Exit Sub
    End If
    If Not IsDate(xDate11.Text) Then
        MsgBox "—«Ã⁄ «· «—ÌŒ 3"
        Exit Sub
    End If
    If Not IsDate(xDate22.Text) Then
        MsgBox "—«Ã⁄ «· «—ÌŒ 4"
        Exit Sub
    End If
    
    With invGrid
    For i = 1 To .Rows - 1
        .TextMatrix(i, 1) = ""
        .TextMatrix(i, 2) = ""
        .TextMatrix(i, 3) = ""
    Next i

'  √Ê· › —…
        cStr1 = " select sum((FILE1_11.[in] - FILE1_11.[out]) * file1_10.cost )  from FILE1_11 inner JOIN file1_10 ON file1_10.item = FILE1_11.item where date < " & DateSq(xDate1.Text)
'        If xMosm1.Text <> "" Then cStr1 = cStr1 & " and mid(file1_10.mosm,1,1) = " & MyParn(xMosm1.Text)
'        If xYear1.Text <> "" Then cStr1 = cStr1 & " and mid(file1_10.mosm,2,2) = " & MyParn(xYear1.Text)
        .TextMatrix(1, 1) = GetDesca(cStr1)
    
        cStr1 = " select sum(sal )  from file4_11 where type = '4' and date >= " & DateSq(xDate1.Text) & " and date <= " & DateSq(xDate2.Text)
'        If xMosm1.Text <> "" Then cStr1 = cStr1 & " and mid(file7_20.mosm,1,1) = " & MyParn(xMosm1.Text)
'        If xYear1.Text <> "" Then cStr1 = cStr1 & " and mid(file7_20.mosm,2,2) = " & MyParn(xYear1.Text)
        .TextMatrix(2, 1) = GetDesca(cStr1)
    
        cStr1 = " select sum(pay )  from file4_11 where type = '5' and date >= " & DateSq(xDate1.Text) & " and date <= " & DateSq(xDate2.Text)
'        If xMosm1.Text <> "" Then cStr1 = cStr1 & " and mid(file6_11.mosm,1,1) = " & MyParn(xMosm1.Text)
'        If xYear1.Text <> "" Then cStr1 = cStr1 & " and mid(file6_11.mosm,2,2) = " & MyParn(xYear1.Text)
        .TextMatrix(3, 1) = GetDesca(cStr1)
        .TextMatrix(4, 1) = Val(.TextMatrix(2, 1)) - Val(.TextMatrix(3, 1))
    
        cStr1 = " select sum(file1_10.cost * out ) as tcost  from Items_Sales inner JOIN file1_10 ON file1_10.item = Items_Sales.item where date >= " & DateSq(xDate1.Text) & " and date <= " & DateSq(xDate2.Text)
'        If xMosm1.Text <> "" Then cStr1 = cStr1 & " and mid(file1_10.mosm,1,1) = " & MyParn(xMosm1.Text)
'        If xYear1.Text <> "" Then cStr1 = cStr1 & " and mid(file1_10.mosm,2,2) = " & MyParn(xYear1.Text)
        .TextMatrix(5, 1) = GetDesca(cStr1)
    
        cStr1 = " select sum(t_total - discount )  from T_SALESDOC where date >= " & DateSq(xDate1.Text) & " and date <= " & DateSq(xDate2.Text)
        .TextMatrix(6, 1) = GetDesca(cStr1)
        .TextMatrix(7, 1) = Val(.TextMatrix(6, 1)) - Val(.TextMatrix(5, 1))
        If Val(.TextMatrix(6, 1)) <> 0 Then .TextMatrix(8, 1) = Val(.TextMatrix(7, 1)) / Val(.TextMatrix(6, 1)) * 100
    
        cStr1 = " select sum((FILE1_11.[in] - FILE1_11.[out]) * file1_10.cost )  from FILE1_11 inner JOIN file1_10 on file1_10.item = FILE1_11.item where date <= " & DateSq(xDate2.Text)
'        If xMosm1.Text <> "" Then cStr1 = cStr1 & " and mid(file1_10.mosm,1,1) = " & MyParn(xMosm1.Text)
'        If xYear1.Text <> "" Then cStr1 = cStr1 & " and mid(file1_10.mosm,2,2) = " & MyParn(xYear1.Text)
        
        .TextMatrix(9, 1) = GetDesca(cStr1)
    
    
'  À«‰Ï › —…
    
        cStr1 = " select sum((FILE1_11.[in] - FILE1_11.[out]) * file1_10.cost )  from FILE1_11 inner JOIN file1_10 on file1_10.item = FILE1_11.item where date < " & DateSq(xDate11.Text)
'        If XMOSM2.Text <> "" Then cStr1 = cStr1 & " and mid(file1_10.mosm,1,1) = " & MyParn(XMOSM2.Text)
'        If xyear2.Text <> "" Then cStr1 = cStr1 & " and mid(file1_10.mosm,2,2) = " & MyParn(xyear2.Text)
        .TextMatrix(1, 2) = GetDesca(cStr1)
    
        cStr1 = " select sum(sal )  from file4_11 where type = '4' and date >= " & DateSq(xDate11.Text) & " and date <= " & DateSq(xDate22.Text)
'        If XMOSM2.Text <> "" Then cStr1 = cStr1 & " and mid(file7_20.mosm,1,1) = " & MyParn(XMOSM2.Text)
'        If xyear2.Text <> "" Then cStr1 = cStr1 & " and mid(file7_20.mosm,2,2) = " & MyParn(xyear2.Text)
        .TextMatrix(2, 2) = GetDesca(cStr1)
    
        cStr1 = " select sum(pay )  from file4_11 where type = '5' and date >= " & DateSq(xDate11.Text) & " and date <= " & DateSq(xDate22.Text)
'        If XMOSM2.Text <> "" Then cStr1 = cStr1 & " and mid(file6_11.mosm,1,1) = " & MyParn(XMOSM2.Text)
'        If xyear2.Text <> "" Then cStr1 = cStr1 & " and mid(file6_11.mosm,2,2) = " & MyParn(xyear2.Text)
        .TextMatrix(3, 2) = GetDesca(cStr1)
        .TextMatrix(4, 2) = Val(.TextMatrix(2, 2)) - Val(.TextMatrix(3, 2))
    
        cStr1 = " select sum(file1_10.cost * out ) as tcost  from Items_Sales inner JOIN file1_10 ON file1_10.item = Items_Sales.item where date >= " & DateSq(xDate11.Text) & " and date <= " & DateSq(xDate22.Text)
'        If XMOSM2.Text <> "" Then cStr1 = cStr1 & " and mid(file1_10.mosm,1,1) = " & MyParn(XMOSM2.Text)
'        If xyear2.Text <> "" Then cStr1 = cStr1 & " and mid(file1_10.mosm,2,2) = " & MyParn(xyear2.Text)
        .TextMatrix(5, 2) = GetDesca(cStr1)
    
        cStr1 = " select sum(t_total - discount )  from T_SALESDOC where date >= " & DateSq(xDate11.Text) & " and date <= " & DateSq(xDate22.Text)
        .TextMatrix(6, 2) = GetDesca(cStr1)
        .TextMatrix(7, 2) = Val(.TextMatrix(6, 2)) - Val(.TextMatrix(5, 2))
        If Val(.TextMatrix(6, 2)) <> 0 Then .TextMatrix(8, 2) = Val(.TextMatrix(7, 2)) / Val(.TextMatrix(6, 2)) * 100
    
        cStr1 = " select sum((FILE1_11.[in] - FILE1_11.[out]) * file1_10.cost )  from FILE1_11 inner JOIN file1_10 on file1_10.item = FILE1_11.item where date <= " & DateSq(xDate22.Text)
'        If XMOSM2.Text <> "" Then cStr1 = cStr1 & " and mid(file1_10.mosm,1,1) = " & MyParn(XMOSM2.Text)
'        If xyear2.Text <> "" Then cStr1 = cStr1 & " and mid(file1_10.mosm,2,2) = " & MyParn(xyear2.Text)
        .TextMatrix(9, 2) = GetDesca(cStr1)
            
        For i = 1 To .Rows - 1
            .TextMatrix(i, 1) = Round(Val(.TextMatrix(i, 1)), 2)
            .TextMatrix(i, 2) = Round(Val(.TextMatrix(i, 2)), 2)
            .TextMatrix(i, 3) = Round(Val(.TextMatrix(i, 2)) - Val(.TextMatrix(i, 1)), 2)
        Next i
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 4
    End With
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
openCon con
With invGrid
    .Cols = 4
    .Rows = 10
    .TextMatrix(1, 0) = "—’Ìœ √Ê·"
    .TextMatrix(2, 0) = "„‘ —Ì« "
    .TextMatrix(3, 0) = "„—œÊœ „‘ —Ì« "
    .TextMatrix(4, 0) = "’«›Ï „‘ —Ì« "
    
    .TextMatrix(5, 0) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(6, 0) = "ﬁÌ„… „»Ì⁄« "
    .TextMatrix(7, 0) = "—»Õ „»Ì⁄« "
    .TextMatrix(8, 0) = "‰”»… —»Õ "
    .TextMatrix(9, 0) = "—’Ìœ √Œ—"
    
    .TextMatrix(0, 1) = "› —… √Ê·Ï"
    .TextMatrix(0, 2) = "› —… À«‰Ì…"
    .TextMatrix(0, 3) = "«·›—ﬁ"
    
    .ColWidth(0) = 2500
    .ColWidth(1) = 3000
    .ColWidth(2) = 3000
    .ColWidth(3) = 2000
    .ColFormat(1) = "#0.00"
    .ColFormat(2) = "#0.00"
    .ColFormat(3) = "#0.00"
End With
End Sub

