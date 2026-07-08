VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form einvoice_menu_oldfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "«·›« Ê—… «·«·Ìﬂ —Ê‰Ì…"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2085
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand cmdSetting 
      Height          =   555
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   979
      _Version        =   196610
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "÷»ÿ «·«⁄œ«œ« "
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin Threed.SSCommand cmdCompany 
      Height          =   555
      Left            =   135
      TabIndex        =   1
      Top             =   720
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   979
      _Version        =   196610
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "÷»ÿ »Ì«‰«  «·‘—ﬂ…"
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   555
      Left            =   135
      TabIndex        =   2
      Top             =   1350
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   979
      _Version        =   196610
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Œ—ÊÃ"
      ButtonStyle     =   3
      PictureAlignment=   9
   End
End
Attribute VB_Name = "einvoice_menu_oldfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCompany_Click()
einvoice_companyfrm.Show 1
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSetting_Click()
einvoice_settingfrm.Show 1
End Sub

