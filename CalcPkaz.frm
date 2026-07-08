VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form CalcPkazFRm 
   Caption         =   " ”⁄Ì— «·«Êﬂ«“ÌÊ‰"
   ClientHeight    =   2910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   4230
      TabIndex        =   16
      Top             =   1665
      Width           =   4560
      Begin VB.CommandButton CMD_DEL2 
         Caption         =   "≈«·€«¡ ”⁄— √Êﬂ«“ÊÌ‰ 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   180
         Width           =   1545
      End
      Begin VB.CommandButton CMD_DEL1 
         Caption         =   "≈·€«¡ ”⁄— «Êﬂ«“ÊÌ‰ 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   180
         Width           =   1545
      End
   End
   Begin VB.CheckBox xIsOkazItem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "«·»Ì⁄ »”⁄— «·«Êﬂ«“ÊÌ‰"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   900
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   540
      Width           =   2040
   End
   Begin VB.TextBox XRATE 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1575
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   90
      Width           =   1365
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -45
      Top             =   315
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
   Begin VB.Frame Frame4 
      Height          =   870
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1665
      Width           =   4065
      Begin VB.CommandButton Command1 
         Caption         =   " ”⁄Ì— «Êﬂ«“ÊÌ‰ 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   180
         Width           =   1365
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
         Height          =   600
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   180
         Width           =   1140
      End
      Begin VB.CommandButton Command2 
         Caption         =   " ”⁄Ì— «Êﬂ«“ÊÌ‰ 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1260
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   180
         Width           =   1365
      End
   End
   Begin MSDataListLib.DataCombo XMOSM 
      Height          =   315
      Left            =   4230
      TabIndex        =   0
      Top             =   90
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
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
   Begin MSDataListLib.DataCombo XFACT 
      Height          =   315
      Left            =   4230
      TabIndex        =   2
      Top             =   495
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
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
   Begin MSDataListLib.DataCombo XGROUP 
      Height          =   315
      Left            =   4230
      TabIndex        =   3
      Top             =   900
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   45
      Top             =   360
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
      Left            =   45
      Top             =   675
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
   Begin ComctlLib.ProgressBar xbar 
      Height          =   240
      Left            =   135
      TabIndex        =   9
      Top             =   2565
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSDataListLib.DataCombo XSECTION 
      Height          =   315
      Left            =   4230
      TabIndex        =   13
      Top             =   1260
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
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
   Begin MSAdodcLib.Adodc DATA5 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·ﬁ”„ :"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   7110
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1260
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰”»… «·Œ’„ :"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   3015
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   135
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·„’‰⁄ :"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   7110
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   495
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·„Ã„Ê⁄… :"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   7155
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   855
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·„Ê”„ :"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   7110
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   135
      Width           =   645
   End
End
Attribute VB_Name = "CalcPkazFRm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection

Private Sub CMD_DEL1_Click()
If MsgBox(" ≈·€«¡ ”⁄— «Êﬂ«“ÊÌ‰ ·ﬂ· «·„ÊœÌ·«  ", vbYesNo + vbDefaultButton2) = vbYes Then
    con.Execute " UPDATE FILE1_10 SET PRICE2 = 0 , OKAZ = 0 ", nRec
    Inform "  „  ⁄œÌ·  " & nRec
End If
End Sub

Private Sub CMD_DEL2_Click()
If MsgBox(" ≈·€«¡ ”⁄— «Êﬂ«“ÊÌ‰ 2 ·ﬂ· «·„ÊœÌ·«  ", vbYesNo + vbDefaultButton2) = vbYes Then
    con.Execute " UPDATE FILE1_10 SET PRICE_2 = 0 , OKAZ_2 = 0 ", nRec
    Inform "  „  ⁄œÌ·  " & nRec
End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub Command2_Click()
Dim cStr1 As String, nCount As Double, I As Double
Dim MyTable As New ADODB.Recordset
If MsgBox(" ”⁄Ì— «Êﬂ«“ÊÌ‰ 2", vbYesNo + vbDefaultButton2) = vbYes Then
    cStr1 = "SELECT * FROM FILE1_10H WHERE MODEL IS NOT NULL "
    If xGroup.BoundText <> "" Then cStr1 = cStr1 & " AND [GROUP] = " & MyParn(xGroup.BoundText)
    If XSECTION.BoundText <> "" Then cStr1 = cStr1 & " AND [SECTION] = " & MyParn(XSECTION.BoundText)
    If xFact.BoundText <> "" Then cStr1 = cStr1 & " AND [FACT] = " & MyParn(xFact.BoundText)
    If xMosm.BoundText <> "" Then cStr1 = cStr1 & " AND [MOSM] = " & MyParn(xMosm.BoundText)
    MyTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
    nCount = MyTable.RecordCount
    If nCount = 0 Then Exit Sub
    XBAR.Visible = True
    XBAR.Min = 0
    XBAR.Max = nCount
    With MyTable
    .MoveFirst
    Do While Not .EOF
        I = I + 1
        XBAR.Value = I
        DoFixPrice !MODELNO, Val(xRate.Text), 2
        .MoveNext
    Loop
    End With
End If
End Sub

Private Sub Command1_Click()
Dim cStr1 As String, nCount As Double, I As Double
Dim MyTable As New ADODB.Recordset
If MsgBox(" ”⁄Ì— «Êﬂ«“ÊÌ‰ 1", vbYesNo + vbDefaultButton2) = vbYes Then
    cStr1 = "SELECT * FROM FILE1_10H WHERE MODEL IS NOT NULL "
    If xGroup.BoundText <> "" Then cStr1 = cStr1 & " AND [GROUP] = " & MyParn(xGroup.BoundText)
    If XSECTION.BoundText <> "" Then cStr1 = cStr1 & " AND [SECTION] = " & MyParn(XSECTION.BoundText)
    If xFact.BoundText <> "" Then cStr1 = cStr1 & " AND [FACT] = " & MyParn(xFact.BoundText)
    If xMosm.BoundText <> "" Then cStr1 = cStr1 & " AND [MOSM] = " & MyParn(xMosm.BoundText)
    MyTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
    nCount = MyTable.RecordCount
    If nCount = 0 Then Exit Sub
    XBAR.Visible = True
    XBAR.Min = 0
    XBAR.Max = nCount
    With MyTable
    .MoveFirst
    Do While Not .EOF
        I = I + 1
        XBAR.Value = I
        DoFixPrice !MODELNO, Val(xRate.Text), 1
        .MoveNext
    Loop
    End With
End If

End Sub

Private Sub Form_Load()
Dim cStr1 As Double
openCon con

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM MOSM ORDER BY DATE DESC "
Set xMosm.RowSource = data1
xMosm.ListField = "DESCA"
xMosm.BoundColumn = "MOSM"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "SELECT * FROM FACT ORDER BY DESCA"
Set xFact.RowSource = DATA2
xFact.ListField = "DESCA"
xFact.BoundColumn = "CODE"

data3.ConnectionString = strCon
data3.RecordSource = "SELECT * FROM FILE1_50 ORDER BY DESCA"
Set xGroup.RowSource = data3
xGroup.ListField = "DESCA"
xGroup.BoundColumn = "CODE"

data5.ConnectionString = strCon
data5.RecordSource = "SELECT * FROM FILE1_10SC ORDER BY DESCA"
Set XSECTION.RowSource = data5
XSECTION.ListField = "DESCA"
XSECTION.BoundColumn = "CODE"
End Sub
Sub DoFixPrice(cModel, nOkaz, SPrice)
    Dim SubItemTable As New ADODB.Recordset
    Dim nRem As Double, nPrice As Double
    Dim nFPrice As Double
    
    Dim cStrUpdate As String
    Dim cStrDEl As String
    
    Dim cStrUpdate_fr As String
    Dim cStrDEl_fr As String

    cStr1 = " SELECT * FROM FILE1_10  WHERE MODELNO = " & MyParn(cModel)
    SubItemTable.Open cStr1, con, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SubItemTable.MoveFirst
        Do While Not SubItemTable.EOF
            If nOkaz > 0 Then
                nPrice = Val(SubItemTable!price & "") * (100 - nOkaz) / 100
                nRem = nPrice - Int(nPrice)
                Select Case nRem
                    Case Is = 0
                        nPrice = Int(nPrice)
                    Case Is <= 0.5
                        nPrice = Int(nPrice) + 0.5
                    Case Is < 1
                        nPrice = Int(nPrice) + 1
                End Select
                If nPrice > SubItemTable!price Then nPrice = SubItemTable!price
            Else
                nPrice = 0
            End If
            If nFPrice = 0 Then nFPrice = nPrice
            
            If SPrice = 1 Then
                If xIsOkazItem.Value <> 0 Then
                    SubItemTable!ISOKAZITEM = 1
                Else
                    SubItemTable!ISOKAZITEM = 0
                End If
                SubItemTable!PRICE2 = nPrice
                SubItemTable!okaz = nOkaz
            Else
                SubItemTable!PRICE_2 = nPrice
                SubItemTable!okaz_2 = nOkaz
            End If
            SubItemTable!IsDate = Date
            
            SubItemTable!IsNEW1 = 1
            SubItemTable!ISNEW2 = 1
            SubItemTable!ISNEW3 = 1
            SubItemTable!ISNEW4 = 1
            SubItemTable!ISNEW5 = 1
            SubItemTable.Update
            
            For sbr = 3 To nCountBranch
                cBr = RetZero(sbr, 2)
                cStrUpdate = cStrUpdate & " INSERT INTO NEW_ITEM (ITEM ,  BRANCH) VALUES        ( " & addvalue(SubItemTable!Item) & " , " & addstring(cBr) & " )  ; "
            Next sbr
            
            For sbr = 61 To nCountBranch_fr
                cBr = RetZero(sbr, 2)
                cStrUpdate_fr = cStrUpdate_fr & " INSERT INTO NEW_ITEM_FR (ITEM ,  BRANCH) VALUES        ( " & addvalue(SubItemTable!Item) & " , " & addstring(cBr) & " )  ; "
            Next sbr
            
            
            SubItemTable.MoveNext
        Loop
        If cStrUpdate <> "" Then con.Execute cStrUpdate
        If cStrUpdate_fr <> "" Then con.Execute cStrUpdate_fr
End Sub
