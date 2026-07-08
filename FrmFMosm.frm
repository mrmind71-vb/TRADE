VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmFMosm 
   Caption         =   "ÑÕíÏ Ãæá ÇáãæÓã"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand CMDNAME 
      Height          =   540
      Left            =   75
      TabIndex        =   6
      Top             =   75
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   953
      _Version        =   196610
      Font3D          =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "1"
   End
   Begin VB.TextBox xBal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2700
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1800
      Width           =   1530
   End
   Begin VB.TextBox xDate 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2700
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1282
      Width           =   1530
   End
   Begin Threed.SSCommand CMD_SAVE 
      Height          =   540
      Left            =   75
      TabIndex        =   2
      Top             =   2400
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   953
      _Version        =   196610
      Font3D          =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ÊÓÌíá ÑÕíÏ Ãæá ÇáãÏÉ"
   End
   Begin MSDataListLib.DataCombo XMOSM 
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   765
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   90
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÑÕíÏ Ãæá ÇáãÏÉ"
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
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1890
      Width           =   1230
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÊÇÑíÎ Ãæá ÇáãÏÉ "
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
      Left            =   4365
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1350
      Width           =   1230
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ãæÓã"
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
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   825
      Width           =   630
   End
End
Attribute VB_Name = "FrmFMosm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CMD_SAVE_Click()
    On Error GoTo myerror
    aRet = aGetDesca("SELECT DATE , F_BAL FROM FILE4_22 WHERE MOSM = " & MyParn(xMosm.BoundText) & " AND CODE = " & MyParn(supMovefrm.xCode.text), con)
    con.BeginTrans
    If UBound(aRet) > 0 Then
        con.Execute " UPDATE FILE4_22 SET DATE = " & addDate(xdate.text) & " , F_BAL = " & Val(xBal.text) & " WHERE CODE = " & MyParn(supMovefrm.xCode.text) & " AND MOSM = " & MyParn(xMosm.BoundText)
    Else
        con.Execute " INSERT INTO FILE4_22 (MOSM,CODE,DATE,F_BAL) " & _
               "Values(" & _
               addstring(xMosm.BoundText) & "," & _
               addstring(supMovefrm.xCode.text) & "," & _
               addDate(xdate.text) & "," & _
               Val(xBal.text) & _
               ")"
    End If
    con.CommitTrans
    Unload Me
    Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Form_Load()
    Me.CMDNAME.Caption = supMovefrm.xDesca.Caption
    con.CursorLocation = adUseClient
    con.Open strCon
    
    data1.ConnectionString = strCon
    data1.RecordSource = "SELECT * FROM MOSM ORDER BY DATE DESC "
    Set xMosm.RowSource = data1
    xMosm.ListField = "DESCA"
    xMosm.BoundColumn = "MOSM"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub xMosm_Validate(Cancel As Boolean)
    aRet = aGetDesca("SELECT DATE , F_BAL FROM FILE4_22 WHERE MOSM = " & MyParn(xMosm.BoundText) & " AND CODE = " & MyParn(supMovefrm.xCode.text), con)
    If UBound(aRet) > 0 Then
        xdate.text = Format(aRet(1), "DD-MM-YYYY")
        xBal.text = Format(aRet(2), "#0.00")
    Else
        xdate.text = ""
        xBal.text = ""
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo) Then SendKeys "{tAB}"
End Sub

