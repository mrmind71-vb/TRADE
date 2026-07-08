VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpClient7 
   Caption         =   " ﬁ«—Ì— «·⁄„·«¡"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
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
   ScaleHeight     =   2670
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   375
      Left            =   1260
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2205
      Width           =   1140
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   375
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2205
      Width           =   1140
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   5250
      Top             =   1725
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   810
      Top             =   -45
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
   Begin VB.Frame Frame1 
      Height          =   2190
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   4890
      Begin VB.TextBox xvalue2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1755
         Width           =   1680
      End
      Begin VB.TextBox xvalue1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1395
         Width           =   1680
      End
      Begin VSFlex7LCtl.VSFlexGrid grid1 
         Height          =   1215
         Left            =   90
         TabIndex        =   0
         Top             =   135
         Width           =   3720
         _cx             =   6562
         _cy             =   2143
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   0
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
         BackColorBkg    =   -2147483633
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
         Rows            =   10
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"rpClient7.frx":0000
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
         Editable        =   2
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰ „Ã„Ê⁄…"
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
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Õ Ï"
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "„‰"
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   780
      End
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
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "rpClient7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If TypeOf ActiveControl Is DBCombo And KeyCode = 46 Then ActiveControl.BoundText = ""
End Sub
Private Sub Form_Load()
openCon con
grid1.ColAlignment(0) = flexAlignRightCenter
grdMake "Select Code,DescA From FILE3_50", "code", "desca", con, grid1
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
doprint1
End Sub
Private Function MYVALID() As Boolean
MYVALID = True
End Function
Private Sub doprint1()
Dim aHeader(1)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset


contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cField1 = myiif("[TYPE] = '4'", _
        " SAL - PAY ") & _
        " as Sales "

cField2 = myiif("", _
        " SAL - PAY ") & _
        " as Balance "

cString = "SELECT File3_11.CODE, File3_10.DESCA, File3_10.[GROUP], FILE3_50.DESCA AS GroupDesca ," & _
          cField1 & "," & _
          cField2 & _
         " FROM (File3_11 INNER JOIN File3_10 ON File3_11.CODE = File3_10.CODE) LEFT JOIN FILE3_50 ON File3_10.[GROUP] = FILE3_50.CODE"

If GrdQry(grid1, "file3_10.[GROUP]", True) <> "" Then
    cString = cString & turnFound(cString) & GrdQry(grid1, "File3_10.[GROUP]", True)
    aHeader(0) = "[" & "„Ã„Ê⁄… ⁄„·«¡ : " & GrdTitle(grid1) & "]"
End If
                 
cString = cString & " GROUP BY File3_11.CODE, File3_10.DESCA, File3_10.[GROUP],FILE3_50.DESCA"
                         
If Val(xvalue1.text) <> 0 Then
    cField = myiif("[TYPE] = '4'", _
        " SAL - PAY ")
    cString = cString & " Having " & cField & " >= " & Val(xvalue1.text)
    aHeader(1) = "[" & BetweenString(IIf(Val(xvalue1.text) = 0, "", xvalue1.text), IIf(Val(xvalue2.text) = 0, "", Val(xvalue2.text)), " ﬁÌ„… „»Ì⁄«  „‰ ", " ≈·Ì ") & "]"
End If

If Val(xvalue2.text) <> 0 Then
    cField = myiif("[TYPE] = '4'", _
        " SAL - PAY ")
    cString = cString & turnFound(cString, " Having ") & cField & " <= " & Val(xvalue2.text)
    aHeader(1) = "[" & BetweenString(IIf(Val(xvalue1.text) = 0, "", xvalue1.text), IIf(Val(xvalue2.text) = 0, "", Val(xvalue2.text)), " ﬁÌ„… „»Ì⁄«  „‰ ", " ≈·Ì ") & "]"
End If
                         
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = sourcetable!CODE
    temptable!str2 = sourcetable![DESCA]
    temptable!str3 = sourcetable![Group]
    temptable!str4 = sourcetable!GroupDesca
    temptable!val3 = sourcetable!balance
    temptable!VAL4 = sourcetable!sales
'    temptable!Date1 = TurnValue(GetDesca("select Max(date) From FILE6_20H where code = " & MyParn(sourcetable!CODE)))
    temptable!str22 = TurnValue(retHeader(aHeader, 0, 2))
    temptable.Update
    sourcetable.MoveNext
Loop
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\CLIENT7.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.Report1.Action = 1
sourcetable.Close
temptable.Close
Set sourcetable = Nothing
Set temptable = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
