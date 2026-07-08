VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rbdoc_trans_brach 
   Caption         =   " ﬁ«—Ì—  ›’Ì·Ï Ê«—œ Ê ’«œ— „ÊœÌ·«  ··›—Ê⁄"
   ClientHeight    =   5550
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
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check5 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬂ· „” ‰œ ’›Õ…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4410
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   5085
      Width           =   1545
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   420
      Left            =   1395
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4995
      Width           =   1275
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4995
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   4830
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   6180
      Begin VB.Frame Frame3 
         Height          =   555
         Left            =   315
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   4185
         Width           =   5820
         Begin VB.OptionButton price 
            Alignment       =   1  'Right Justify
            Caption         =   "„” Â·ﬂ"
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
            Left            =   3570
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   180
            Width           =   960
         End
         Begin VB.OptionButton price 
            Alignment       =   1  'Right Justify
            Caption         =   "«· ﬂ·›…"
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
            Left            =   4770
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   180
            Value           =   -1  'True
            Width           =   960
         End
      End
      Begin VB.Frame Frame2 
         Height          =   600
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   3600
         Width           =   6090
         Begin VB.CheckBox Check4 
            Alignment       =   1  'Right Justify
            Caption         =   " ÕÊÌ·«  „‰"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   135
            Value           =   1  'Checked
            Width           =   1230
         End
         Begin VB.CheckBox Check3 
            Alignment       =   1  'Right Justify
            Caption         =   " ÕÊÌ·«  ≈·Ï"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1530
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   135
            Value           =   1  'Checked
            Width           =   1230
         End
         Begin VB.CheckBox Check2 
            Alignment       =   1  'Right Justify
            Caption         =   "„—œÊœ „‘ —Ì« "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3015
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   135
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "„‘ —Ì« "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4725
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   135
            Value           =   1  'Checked
            Width           =   1230
         End
      End
      Begin MSDataListLib.DataCombo xmosm 
         Height          =   315
         Left            =   1305
         TabIndex        =   5
         Top             =   1452
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xstore1 
         Height          =   315
         Left            =   1305
         TabIndex        =   7
         Top             =   2250
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xfact 
         Height          =   315
         Left            =   1305
         TabIndex        =   10
         Top             =   225
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XMOSM2 
         Height          =   315
         Left            =   1305
         TabIndex        =   17
         Top             =   1851
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   315
         Left            =   1305
         TabIndex        =   18
         Top             =   2655
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker xdate1 
         Height          =   375
         Left            =   2970
         TabIndex        =   24
         Top             =   585
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   661
         _Version        =   393216
         MousePointer    =   5
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
         Format          =   161480705
         CurrentDate     =   43097
      End
      Begin MSComCtl2.DTPicker xdate2 
         Height          =   375
         Left            =   2970
         TabIndex        =   25
         Top             =   990
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   661
         _Version        =   393216
         MousePointer    =   5
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
         Format          =   161480705
         CurrentDate     =   43097
      End
      Begin MSDataListLib.DataCombo xsection 
         Height          =   315
         Left            =   1305
         TabIndex        =   26
         Top             =   3060
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label8 
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
         Height          =   255
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   3105
         Width           =   555
      End
      Begin VB.Label Label7 
         Caption         =   "«·„Ê—œ :"
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   2700
         Width           =   1320
      End
      Begin VB.Label Label4 
         Caption         =   "„Ê”„ «·„” ‰œ"
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1890
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "„’‰⁄ :"
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
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label Label6 
         Caption         =   "«·›—⁄ :"
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
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   2340
         Width           =   1005
      End
      Begin VB.Label Label5 
         Caption         =   "„Ê”„ «·„ÊœÌ·"
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
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1446
         Width           =   1140
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Õ Ì :"
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
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1059
         Width           =   465
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Height          =   195
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   672
         Width           =   765
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc DATA5 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc data6 
      Height          =   330
      Left            =   0
      Top             =   0
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
End
Attribute VB_Name = "rbdoc_trans_brach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearchFact As New Search3
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


data5.ConnectionString = strCon
data5.RecordSource = "Select Code,DescA From File4_10 order by Desca"
Set xSupp.RowSource = data5
xSupp.ListField = "Desca"
xSupp.BoundColumn = "Code"

DATA6.ConnectionString = strCon
DATA6.RecordSource = "Select code,desca From file1_10sc ORDER BY code "
Set xSection.RowSource = DATA6
xSection.ListField = "Desca"
xSection.BoundColumn = "Code"


data4.ConnectionString = strCon
data4.RecordSource = "Select MOSM,DescA From MOSM     ORDER by DATE "
Set xMosm.RowSource = data4
xMosm.ListField = "Desca"
xMosm.BoundColumn = "MOSM"

Set xmosm2.RowSource = data4
xmosm2.ListField = "Desca"
xmosm2.BoundColumn = "MOSM"

data3.ConnectionString = strCon
data3.RecordSource = "Select code , DescA From fact ORDER by desca "
Set xFact.RowSource = data3
xFact.ListField = "Desca"
xFact.BoundColumn = "code"

data1.ConnectionString = strCon
data1.RecordSource = "Select code ,DescA From file0_40 ORDER by code "
Set xStore1.RowSource = data1
xStore1.ListField = "Desca"
xStore1.BoundColumn = "code"

If cBranch <> "00" Then
    xStore1.BoundText = cBranchStore
    xStore1.Enabled = False
End If


End Sub
Private Sub CmdApply_Click()
    If Not MYVALID Then Exit Sub
    doprint1
End Sub
Private Function MYVALID() As Boolean

If Not IsDate(xdate2.Value) And Trim(xdate2.Value) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ”·Ì„"
    Exit Function
End If

'If xstore1.BoundText = "" Then
'    MsgBox "«Œ Ì«— «·›—⁄ "
'    Exit Function
'End If
'
MYVALID = True
End Function
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub xFACT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    FactLookupAll Me, oSearchFact
End If
End Sub
Sub myProc()
    xFact.BoundText = oSearchFact.grid1.TextMatrix(oSearchFact.grid1.Row, 0)
    Unload oSearchFact
End Sub
Private Sub suplookup()
    Dim Generalarray(5)
    Dim listarray(0, 4)
    Dim GrdArray(1, 1)
    
    Set Generalarray(0) = Me
    Generalarray(1) = "Select Code, DescA From FILE4_10"
    Generalarray(2) = "Order by file4_10.Desca"
    Generalarray(3) = 4200
    Generalarray(5) = False
    
    listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
    listarray(0, 1) = "(%%DESCA%%) "
    
    GrdArray(0, 0) = "ﬂÊœ «·„Ê—œ"
    GrdArray(0, 1) = 1000
    
    GrdArray(1, 0) = "≈”„ «·„Ê—œ"
    GrdArray(1, 1) = 3000
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Load Search3
    Search3.Caption = "«” ⁄·«„"
    Search3.Show 1
End Sub
Private Sub doprint1()
Dim aHeader(5)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

If Trim(xFact.BoundText) <> "" Then
    cWhere = cWhere & " AND Q_ALL_TRANS.FACT = " & MyParn(xFact.BoundText)
    aHeader(1) = "[" & "„’‰⁄ : " & xFact.Text & "]"
End If
          
If IsDate(xDate1.Value) Then
    cWhere = cWhere & " AND date >= " & DateSq(xDate1.Value)
    aHeader(2) = "[" & BetweenString(xDate1.Value & "", xdate2.Value & "") & "]"
End If

If IsDate(xdate2.Value) Then
    cWhere = cWhere & " AND date <= " & DateSq(xdate2.Value)
    aHeader(2) = "[" & BetweenString(xDate1.Value & "", xdate2.Value & "") & "]"
End If

If Trim(xMosm.BoundText) <> "" Then
    cWhere = cWhere & " AND Q_ALL_TRANS.mosm  = " & MyParn(xMosm.BoundText)
    aHeader(3) = "[" & "„Ê”„ : " & xMosm.Text & "]"
End If

If Trim(xmosm2.BoundText) <> "" Then
    cWhere = cWhere & " AND Q_ALL_TRANS.mosm2 = " & MyParn(xmosm2.BoundText)
    aHeader(3) = "[" & "„Ê”„ «·„” ‰œ : " & xmosm2.Text & "]"
End If

If Trim(xStore1.BoundText) <> "" Then
    cWhere = cWhere & " AND store = " & MyParn(xStore1.BoundText)
End If

If Trim(xSection.BoundText) <> "" Then
    cWhere = cWhere & " AND Q_ALL_TRANS.[section] = " & Val(xSection.BoundText)
    aHeader(4) = "[" & "ﬁ”„ : " & xSection.Text & "]"
End If

If Trim(xSupp.BoundText) <> "" Then
    cWhere = cWhere & " AND Q_ALL_TRANS.SUPPLER = " & MyParn(xSupp.BoundText)
    aHeader(5) = "[" & "«·„Ê—œ : " & xSupp.Text & "]"
End If

CWT = ""
If Check1.Value <> 0 Then CWT = CWT & " OR [TYPE] = 1 "
If Check2.Value <> 0 Then CWT = CWT & " OR [TYPE] = 3 "
If Check3.Value <> 0 Then CWT = CWT & " OR [TYPE] = 2 "
If Check4.Value <> 0 Then CWT = CWT & " OR [TYPE] = 4 "

If CWT = "" Then
    Exit Sub
Else
    CWT = " AND ( " & Mid(CWT, 4) & " ) "
End If
cString = "SELECT Q_ALL_TRANS.*  , FILE1_10.PRICE AS ITEMPRICE , file0_40.desca as storedesca from Q_ALL_TRANS INNER JOIN FILE1_10 ON FILE1_10.ITEM = Q_ALL_TRANS.ITEM  inner join file0_40 on file0_40.code = q_all_trans.store  WHERE DOC_NO IS NOT NULL " & cWhere & CWT
                                                     
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = ArbString(sourcetable!doc_no & " ( " & sourcetable!INV_NO & " ) ")
    temptable!str2 = ArbString(sourcetable!TypeDesca & "  " & TurnValue(sourcetable!MOSM2, Null, ""))
    temptable!STR11 = sourcetable!typedoc
    temptable!Date1 = sourcetable!Date
    
    temptable!str4 = TurnValue(sourcetable!FACTDESCA, "", Null)
    temptable!str6 = TurnValue(DelZero(sourcetable!modelfact), "", Null)
    temptable!str16 = TurnValue((sourcetable!modelfact), "", Null)
    temptable!str5 = TurnValue(sourcetable!DESCA, "", Null)
    temptable!STR7 = TurnValue(sourcetable!MOSM, "", Null)
    temptable!str9 = TurnValue(sourcetable!SCAL, "", Null)
    temptable!str19 = TurnValue(sourcetable!storedesca, "", Null)
    temptable!val17 = TurnValue(sourcetable!C_SCAL, "", Null)
    temptable!str10 = TurnValue(sourcetable!MODELNO, "", Null)
    
    temptable!val2 = Val(sourcetable!Q_IN & "")
    temptable!VAL4 = Val(sourcetable!Q_OUT & "")
    
    If price(0).Value Then
        temptable!VAL1 = sourcetable!cost_inv
        temptable!val3 = sourcetable!cost_inv * (sourcetable!Q_IN - sourcetable!Q_OUT)
    End If
    
    If price(1).Value Then
        temptable!VAL1 = sourcetable!itemPrice
        temptable!val3 = Val(sourcetable!itemPrice & "") * (Val(sourcetable!Q_IN & "") - Val(sourcetable!Q_OUT & ""))
    End If
    
    
    temptable!str8 = " ›’Ì·Ï Ê«—œ Ê ’«œ— „ÊœÌ·«  ·›—⁄ " & xStore1.Text
    temptable!str21 = TurnValue(retHeader(aHeader, 0, 6))
    temptable.Update
    sourcetable.MoveNext
Loop
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
If Check5.Value Then
    main.REPORT1.ReportFileName = App.Path & "\Reports\ro_all_transp.rpt"
Else
    main.REPORT1.ReportFileName = App.Path & "\Reports\ro_all_trans.rpt"
End If
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
sourcetable.Close
temptable.Close
Set sourcetable = Nothing
Set temptable = Nothing
End Sub


