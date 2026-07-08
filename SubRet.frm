VERSION 5.00
Begin VB.Form SubRet 
   Caption         =   "ÈíÇäÇÊ ÇáÔÍä"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6690
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2655
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ÍÝÙ "
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4725
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2100
      Width           =   1815
   End
   Begin VB.CommandButton CmdUndo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ÊÑÇÌÚ"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   225
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2100
      Width           =   1815
   End
   Begin VB.TextBox xRem4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   75
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1500
      Width           =   5160
   End
   Begin VB.TextBox xRem1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   75
      MaxLength       =   100
      TabIndex        =   0
      Top             =   225
      Width           =   5160
   End
   Begin VB.TextBox xRem2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3825
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   650
      Width           =   1410
   End
   Begin VB.TextBox xRem3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   75
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1075
      Width           =   5160
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ÅÓã ÇáãÈáÛ ÚäÉ"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5385
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1515
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ÑÞã ÇáÈæáíÕÉ"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5385
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1095
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ÔÑßÉ ÇáÔÍä"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5385
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   255
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ÇáÊÇÑíÎ"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5370
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   675
      Width           =   495
   End
End
Attribute VB_Name = "SubRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim SubRetTable As New ADODB.Recordset
Private Sub cmdSave_Click()
openCon con

With SubRetTable
    If GetDesca("select doc_no from  sub_ret where doc_no = " & MyParn(Purchasefrm.xdoc_no.text), con) = "" Then
        con.Execute " INSERT INTO sub_ret (doc_no,rem1 , rem2 , rem3 , rem4 ) " & _
               "Values(" & _
               addstring(Purchasefrm.xdoc_no.text) & "," & _
               addstring(xRem1.text) & "," & _
               addstring(xRem2.text) & "," & _
               addstring(xRem3.text) & "," & _
               addstring(xRem4.text) & _
               ")"
    Else
        con.Execute " update sub_ret SET " & _
               " rem1 = " & addstring(xRem1.text) & "," & _
               " rem2 = " & addstring(xRem2.text) & "," & _
               " rem3 = " & addstring(xRem3.text) & "," & _
               " rem4 = " & addstring(xRem4.text) & _
               " where doc_no = " & MyParn(Purchasefrm.xdoc_no.text)
    End If
End With
con.Close
Set con = Nothing
Unload Me
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub Form_Load()
aRet = aGetDesca("select rem1 , rem2 , rem3 , rem4 FROM SUB_RET WHERE DOC_NO = " & MyParn(Purchasefrm.xdoc_no.text), con)
If UBound(aRet) > 0 Then
    xRem1.text = aRet(1) & ""
    xRem2.text = aRet(2) & ""
    xRem3.text = aRet(3) & ""
    xRem4.text = aRet(4) & ""
End If
End Sub
