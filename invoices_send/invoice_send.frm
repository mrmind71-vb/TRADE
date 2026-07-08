VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form invoice_sendfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "›Ê« Ì— «—”«· »Ì«‰«  «·›—⁄"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   3285
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
      Caption         =   "÷»ÿ «⁄œ«œ«  «—”«· »Ì«‰«  «·›—⁄"
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   555
      Left            =   135
      TabIndex        =   1
      Top             =   1260
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
   Begin Threed.SSCommand cmdSendInvoices 
      Height          =   555
      Left            =   135
      TabIndex        =   2
      Top             =   675
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
      Caption         =   "√—”«· ›Ê« Ì— „»Ì⁄«  €Ì— „—”·…"
      ButtonStyle     =   3
      PictureAlignment=   9
   End
End
Attribute VB_Name = "invoice_sendfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdSendInvoices_Click()
InvoiceTransfrm.Show 1
End Sub
Private Sub cmdSetting_Click()
invoice_send_settingfrm.Show 1
End Sub

