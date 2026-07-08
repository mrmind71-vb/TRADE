VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpitem9 
   Caption         =   " ﬁ«—Ì— "
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
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
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Report1 
      Left            =   6615
      Top             =   2475
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   3285
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   5865
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "„»Ì⁄«  ›ﬁÿ"
         Height          =   195
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2925
         Width           =   1455
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   2790
         Width           =   1635
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   2430
         Width           =   1635
      End
      Begin MSDataListLib.DataCombo xstore 
         Height          =   315
         Left            =   900
         TabIndex        =   1
         Top             =   1710
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grid1 
         Height          =   1440
         Left            =   540
         TabIndex        =   0
         Top             =   195
         Width           =   3765
         _cx             =   6641
         _cy             =   2540
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
         FormatString    =   $"rpitem9.frx":0000
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
      Begin MSDataListLib.DataCombo XUNIT 
         Height          =   315
         Left            =   900
         TabIndex        =   9
         Top             =   2070
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·Ì :"
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
         Index           =   2
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   2835
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„‰ :"
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
         Index           =   1
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2475
         Width           =   330
      End
      Begin VB.Label Label3 
         Caption         =   "«·ÊÕœ…"
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
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2115
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "„Ã„Ê⁄… —∆Ì”Ì…"
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
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label4 
         Caption         =   "„Œ“‰"
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
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1755
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   " ›—Ì€"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2565
      RightToLeft     =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3330
      Width           =   1140
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "«” Ã«»…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3735
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3330
      Width           =   1185
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1215
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3330
      Width           =   1320
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   45
      Top             =   1125
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   180
      Top             =   630
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
   Begin VB.Label Label6 
      Height          =   255
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2175
      Width           =   1005
   End
End
Attribute VB_Name = "rpitem9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdApply_Click()
doprint1
End Sub
Private Sub CmdClear_Click()
xStore.BoundText = ""
XUNIT.BoundText = ""
grid1.Rows = 0
grid1.Rows = 10
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
grdMake "Select Code,DescA From File1_51", "code", "desca", CON, grid1

data1.ConnectionString = CON.ConnectionString
data1.RecordSource = "Select Code,DescA From File0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

DATA2.ConnectionString = CON.ConnectionString
DATA2.RecordSource = "Select Code,DescA From File1_13 order by code"
Set XUNIT.RowSource = DATA2
XUNIT.ListField = "Desca"
XUNIT.BoundColumn = "Code"
End Sub
Private Sub doprint1()
Dim temptable As New ADODB.Recordset
Dim sourcetable As ADODB.Recordset
Dim aHeader(4), nFirstbalance As Double
contemp.Execute "delete * from temp"

temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

If IsDate(xdate1.Text) Then
cString = "Select Sum(val([in] & '')- val([out] & '')) From (file1_11 inner join file1_10 on file1_11.item = file1_10.item) inner join file1_50 on file1_10.group = file1_50.code Where " & _
                    "[Date] < " & DateSq(xdate1.Text)
    
    If GrdQry(grid1, "file1_50.Group", True) <> "" Then
        cString = cString & turnFound(cString) & GrdQry(grid1, "File1_50.Group", True)
        aHeader(0) = "[" & "«·„Ã„Ê⁄… : " & GrdTitle(grid1) & "]"
    End If
    
    If xStore.BoundText <> "" Then
        cString = cString & turnFound(cString) & "File1_11.store = " & MyParn(xStore.BoundText)
        aHeader(1) = "[" & "„Œ“‰ : " & xStore.Text & "]"
    End If
    
    If XUNIT.BoundText <> "" Then
        cString = cString & turnFound(cString) & "File1_10.UNIT = " & MyParn(XUNIT.BoundText)
        aHeader(2) = "[" & "«·ÊÕœ… : " & XUNIT.Text & "]"
    End If
    nFirstbalance = Val(GetDesca(cString) & "")
End If

cField1 = myiif(" ([Type] = '2' OR [Type] = '20')" _
          , " val( [in] & '')") & _
           " As PURCHASE "

cField2 = myiif(" ([Type] = '5' OR [Type] = '23')" _
          , "val( [in] & '')") & _
           " As RETSALES "

cField3 = myiif(" ([Type] = '8' )" _
          , "val( [in] & '')") & _
           " As Trans_in "

cField4 = myiif(" ([Type] = '10' )" _
          , "val( [in] & '')") & _
           " As [INPUT]"

cField5 = myiif(" ([Type] = '1' )" _
          , "val( [in] & '')") & _
           " As Stock"

cField6 = myiif(" ([Type] = '3' OR [Type] = '21')" _
          , "val( [out] & '')") & _
           " As SALES "

cField7 = myiif(" ([Type] = '4' OR [Type] = '22')" _
          , "val( [out] & '')") & _
           " As RETPUR "

cField8 = myiif(" ([Type] = '7' )" _
          , "val( [out] & '')") & _
           " As Trans_out "

cField9 = myiif(" ([Type] = '9' )" _
          , "val( [out] & '')") & _
           " As [OUTPUT] "

cField10 = myiif(" ([Type] = '6' )" _
          , "val( [out] & '')") & _
           " As Damage "

cField11 = myiif("", "val([in] & '')- val([out] & '')") & _
          " As TotalDay"

cString = "SELECT FILE1_11.DATE," & _
           cField1 & "," & cField2 & "," & cField3 & "," & _
           cField4 & "," & cField5 & "," & cField6 & "," & _
           cField7 & "," & cField8 & "," & cField9 & "," & _
           cField10 & "," & cField11 & _
          " FROM (FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM)INNER JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE "


If GrdQry(grid1, "file1_50.Group", True) <> "" Then
    cString = cString & turnFound(cString) & GrdQry(grid1, "File1_50.Group", True)
    aHeader(0) = "[" & "«·„Ã„Ê⁄… : " & GrdTitle(grid1) & "]"
End If

If xStore.BoundText <> "" Then
    cString = cString & turnFound(cString) & "File1_11.store = " & MyParn(xStore.BoundText)
    aHeader(1) = "[" & "„Œ“‰ : " & xStore.Text & "]"
End If

If XUNIT.BoundText <> "" Then
    cString = cString & turnFound(cString) & "File1_10.UNIT = " & MyParn(XUNIT.BoundText)
    aHeader(2) = "[" & "«·ÊÕœ… : " & XUNIT.Text & "]"
End If

If IsDate(xdate1.Text) Then
    cString = cString & turnFound(cString) & "date >= " & DateSql(xdate1.Text)
    aHeader(3) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

If IsDate(xDate2.Text) Then
    cString = cString & turnFound(cString) & "date <= " & DateSql(xDate2.Text)
    aHeader(3) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

cString = cString & " GROUP BY FILE1_11.date"

If (Check1.Value <> 0) Then
    cString = cString & " Having " & myiif(" ([Type] = '3' OR [Type] = '23')" _
          , "val( [out] & '')") & " <> 0"

End If
Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, CON, adOpenForwardOnly, adLockReadOnly, adCmdText

With sourcetable
    Do Until .EOF
        temptable.AddNew
        temptable!Date1 = sourcetable!Date
        temptable!val1 = nFirstbalance
        temptable!val2 = sourcetable!purchase
        temptable!Val3 = sourcetable!Retsales
        temptable!val4 = sourcetable!Trans_in
        temptable!val5 = sourcetable!Input
        temptable!Val6 = sourcetable!Stock
        If Check1.Value = 0 Then
            temptable!val7 = sourcetable!Sales
        Else
            temptable!val7 = sourcetable!Sales - sourcetable!Retsales
        End If
        temptable!val8 = sourcetable!Retpur
        temptable!Val9 = sourcetable!Trans_out
        temptable!Val10 = sourcetable!output
        temptable!Val11 = sourcetable!Damage
        temptable!Val12 = sourcetable!TotalDay
        temptable!Val13 = nFirstbalance + sourcetable!TotalDay
        nFirstbalance = nFirstbalance + sourcetable!TotalDay
        temptable!str21 = TurnValue(retHeader(aHeader, 0, 4))
        temptable.Update
        .MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    REPORT1.ReportFileName = IIf(Check1.Value = 0, App.Path & "\Reports\Item9.rpt", App.Path & "\Reports\Item91.rpt")
    contemp.BeginTrans
    contemp.CommitTrans
    REPORT1.DataFiles(0) = "c:\elmorshed\temp.mdb"
    REPORT1.Action = 1
End If

temptable.Close
sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xdate1.Text) And Trim(xdate1.Text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’ÕÌÕ"
    Exit Function
End If
If Not IsDate(xDate2.Text) And Trim(xDate2.Text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’ÕÌÕ"
    Exit Function
End If
MYVALID = True
End Function


