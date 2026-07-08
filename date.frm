VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form datefrm 
   ClientHeight    =   3135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4395
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4395
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.MonthView xdate 
      Height          =   3060
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   5398
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   159186951
      CurrentDate     =   41703
   End
End
Attribute VB_Name = "datefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oDate As TextBox
Private Sub Form_Load()
If IsDate(oDate.Text) Then
    xDate.Value = Format(oDate, "YYYY-MM-DD")
End If
End Sub
Private Sub xdate_DateClick(ByVal DateClicked As Date)
If IsDate(DateClicked) Then
    oDate.Text = Format(DateClicked, "DD-MM-YYYY")
    Unload Me
End If
End Sub
