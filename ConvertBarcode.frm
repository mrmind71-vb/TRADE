VERSION 5.00
Begin VB.Form ConvertBarcode 
   Caption         =   " ⁄œÌ· »«—ﬂÊœ ··„Õ·« "
   ClientHeight    =   2445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2445
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2040
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   180
      Width           =   4200
      Begin VB.CommandButton CMD_OK 
         Caption         =   " ⁄œÌ·"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   675
         TabIndex        =   5
         Top             =   1305
         Width           =   2850
      End
      Begin VB.TextBox XITEMNEW 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   675
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   765
         Width           =   1545
      End
      Begin VB.TextBox XITEMOLD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   675
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "»«—ﬂÊœ ’ÕÌÕ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   855
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "»«—ﬂÊœ „Õ–Ê› :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   315
         Width           =   1290
      End
   End
End
Attribute VB_Name = "ConvertBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection

Private Sub CMD_OK_Click()
If Val(XITEMNEW.text) > 100 And Val(XITEMOLD.text) > 100 Then
    If MsgBox(" ⁄œÌ· «·»«—ﬂÊœ «·„Õ–Ê› ≈·Ï »«—ﬂÊœ «·’ÕÌÕ", vbYesNo + vbDefaultButton2) = vbYes Then
        con.Execute " UPDATE FILE6_20 SET ITEM = " & Val(XITEMNEW.text) & " WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
        Inform "  ⁄œÌ· ⁄œœ „»Ì⁄«  " & nRec
    
        con.Execute " UPDATE FILE7_20 SET ITEM = " & Val(XITEMNEW.text) & " WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
        Inform "  ⁄œÌ· ⁄œœ „‘ —Ì« " & nRec
    
        con.Execute " UPDATE FILE7_10 SET ITEM = " & Val(XITEMNEW.text) & " WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
        Inform "  ⁄œÌ· ⁄œœ „— Ã⁄  " & nRec
    
        con.Execute " UPDATE FILE1_60 SET ITEM = " & Val(XITEMNEW.text) & " WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
        con.Execute " UPDATE FILE1_61 SET ITEM = " & Val(XITEMNEW.text) & " WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
        Inform "  ⁄œÌ· ⁄œœ „— Ã⁄  " & nRec
    
        If cBranch = "00" Then
            con.Execute " UPDATE FILE6_50 SET ITEM = " & Val(XITEMNEW.text) & " WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
            Inform "  ⁄œÌ· ⁄œœ ÿ·»Ì«  ‘—«¡" & nRec
    
            con.Execute " UPDATE FILE6_51 SET ITEM = " & Val(XITEMNEW.text) & " WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
            Inform "  ⁄œÌ· ⁄œœ ÿ·»Ì«  ⁄„·«¡ " & nRec
    
            con.Execute " UPDATE FILE6_52 SET ITEM = " & Val(XITEMNEW.text) & " WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
            Inform "  ⁄œÌ· ⁄œœ ≈” ·«„ ÿ·»Ì«  ⁄„·«¡ " & nRec
    
            con.Execute " UPDATE FILE6_55 SET ITEM = " & Val(XITEMNEW.text) & " WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
            Inform "  ⁄œÌ· ⁄œœ ≈” ·«„ ÿ·»Ì«  ‘—«¡ " & nRec
        End If
    
        con.Execute " UPDATE FILE0_10  SET ITEM = " & Val(XITEMNEW.text) & " WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
        con.Execute " UPDATE FILE0_100 SET ITEM = " & Val(XITEMNEW.text) & " WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
        Inform "  ⁄œÌ· ⁄œœ Ã—œ " & nRec
        If cBranch = "00" Then
            con.Execute " UPDATE FR6_20 SET ITEM = " & Val(XITEMNEW.text) & " WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
            con.Execute " UPDATE FR7_20 SET ITEM = " & Val(XITEMNEW.text) & " WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
            con.Execute " UPDATE FR7_10 SET ITEM = " & Val(XITEMNEW.text) & " WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
            con.Execute " UPDATE FR1_60 SET ITEM = " & Val(XITEMNEW.text) & " WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
            con.Execute " UPDATE FR0_10  SET ITEM = " & Val(XITEMNEW.text) & " WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
            con.Execute " UPDATE FR0_100 SET ITEM = " & Val(XITEMNEW.text) & " WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
        End If
        
        con.Execute " DELETE FROM FILE1_10  WHERE ITEM = " & MyParn(XITEMOLD.text), nRec
        Inform "  „ «· ⁄œÌ· ≈·Ï «·»«—ﬂÊœ «·’ÕÌÕ Ê Õ–› «·»«—ﬂÊœ «·„·€Ï"
    End If
End If
End Sub

Private Sub Form_Load()
    
    con.CursorLocation = adUseClient
    con.Open strCon
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
