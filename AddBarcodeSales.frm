VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form AddBarcodeSales 
   Caption         =   " ›—Ì€ ÿ·»Ì… „‰ «·»«—ﬂÊœ"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   3105
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar BAR1 
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Top             =   6840
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton CMD_ADD 
      Caption         =   "≈÷«›… «·„ÊœÌ·«  ›Ï «·›« Ê—…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7200
      Width           =   2580
   End
   Begin VB.TextBox XBARCODE 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6585
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "AddBarcodeSales.frx":0000
      Top             =   180
      Width           =   2580
   End
End
Attribute VB_Name = "AddBarcodeSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CMD_ADD_Click()
    Dim cItem As String
    Dim cMOdel As String, nRec As Double
    aLocal = Split(XBARCODE.Text, "00")
    BAR1.Max = UBound(aLocal)
    BAR1.Min = 0
    For i = 1 To UBound(aLocal)
        cItem = Trim(aLocal(i))
        cMOdel = GetDesca("SELECT MODEL FROM FILE1_10 WHERE ITEM = " & Val(cItem))
        con.Execute "  DELETE FROM FILE6_20 WHERE DOC_NO = " & MyParn(Sales_Gom.xDoc_No.Text) & " AND MODEL = " & MyParn(cMOdel)
        con.Execute "  INSERT INTO FILE6_20(DOC_NO , ITEM, MODEL, PRICE, QUANT) SELECT     " & addstring(Sales_Gom.xDoc_No.Text) & " ,  ITEM, model, cost2, 1  FROM         FILE1_10    WHERE MODEL = " & MyParn(cMOdel), nRec
        BAR1.Value = i
    Next i
    Unload Me
End Sub
Private Sub Form_Load()
    openCon con
    XBARCODE.Text = ""
    BAR1.Value = 0
End Sub
