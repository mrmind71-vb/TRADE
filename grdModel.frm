VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form grdModelfrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13860
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4995
   ScaleWidth      =   13860
   Begin VB.Frame Frame1 
      Height          =   2220
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   45
      Width           =   3435
      Begin VB.TextBox XITEM2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   465
         Left            =   270
         MaxLength       =   40
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   765
         Width           =   1905
      End
      Begin VB.TextBox xItem1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   465
         Left            =   270
         MaxLength       =   40
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   225
         Width           =   1905
      End
      Begin VB.CommandButton CMD_BARCODECHANGE 
         Caption         =   " ⁄œÌ· »«—ﬂÊœ "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1350
         Width           =   2625
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "»«—ﬂÊœ ÃœÌœ"
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
         Left            =   2340
         TabIndex        =   10
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "»«—ﬂÊœ „·€Ï"
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
         Left            =   2385
         TabIndex        =   9
         Top             =   810
         Width           =   885
      End
   End
   Begin VB.CommandButton CMD_FIX 
      Caption         =   "FIX BARCODE"
      Height          =   600
      Left            =   2205
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3510
      Width           =   1320
   End
   Begin VB.CommandButton cmdBarCode 
      Height          =   600
      Left            =   2205
      Picture         =   "grdModel.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4185
      Width           =   1320
   End
   Begin VB.Frame Frame4 
      Caption         =   "„ﬁ«”«  Ê«·Ê«‰"
      Height          =   2940
      Left            =   3510
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   45
      Width           =   10185
      Begin VSFlex7Ctl.VSFlexGrid grdModel 
         Height          =   2550
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   9915
         _cx             =   17489
         _cy             =   4498
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   4
         FixedCols       =   2
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   4590
      Top             =   -90
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
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   1740
      Left            =   3555
      TabIndex        =   2
      Top             =   3060
      Width           =   10050
      _cx             =   17727
      _cy             =   3069
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc data20 
      Height          =   330
      Left            =   3960
      Top             =   -45
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
End
Attribute VB_Name = "grdModelfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public pString As String, pCaption As String, pFieldValue As String, pModelNo As String
Private Sub CMR_PRINT_Click()
    Load PrintGrd
    PrintGrd.doprint grid1, 1, , Me.Caption
    PrintGrd.Show 1
End Sub

Private Sub CMD_BARCODECHANGE_Click()
Dim nBaR1 As Double
Dim nBar2 As Double

nBaR1 = Val(xItem1.text)
nBar2 = Val(XITEM2.text)
If nBaR1 = 0 Or nBar2 = 0 Then Exit Sub
With grid2
    If MsgBox("Ÿ»ÿ »«—ﬂÊœ „ﬂ——", vbYesNo + vbDefaultButton2) = vbYes Then
        MsgBox "ÌÃ» ÿ»«⁄… »«—ﬂÊœ ··«’‰«› «· Ï ·Â« »«—ﬂÊœ „ﬂ—— Œÿ√"
        If nBaR1 > 0 And nBar2 > 0 Then
            con.Execute " UPDATE FILE6_20       SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            con.Execute " UPDATE FILE7_10       SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            con.Execute " UPDATE FILE7_20       SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            con.Execute " UPDATE FILE1_60       SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            con.Execute " UPDATE FILE1_61       SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            con.Execute " UPDATE FILE0_10       SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            con.Execute " UPDATE FILE0_100      SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            
            con.Execute " UPDATE FILE6_50       SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            con.Execute " UPDATE FILE6_51       SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            con.Execute " UPDATE FILE6_52       SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            con.Execute " UPDATE FILE6_55       SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            
            con.Execute " UPDATE FILE1_11_ALL   SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            
            con.Execute " UPDATE FR7_20         SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            con.Execute " UPDATE FR1_60         SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            con.Execute " UPDATE FR7_10         SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            con.Execute " UPDATE FR6_20         SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            con.Execute " UPDATE FR0_10         SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            con.Execute " UPDATE FR0_100        SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
            
            cString = "INSERT INTO BARCODE_EDIT (ITEMNEW ,ITEMOLD ) VALUES (  " & nBaR1 & " , " & nBar2 & " ) "
            con.Execute cString
            
            cStrUpdate = ""
            For sbr = 3 To nCountBranch
                cBr = RetZero(sbr, 2)
                cStrUpdate = cStrUpdate & " INSERT INTO NEW_ITEM (ITEM ,  BRANCH) VALUES        ( " & addvalue(nBaR1) & " , " & addstring(cBr) & " )  ; "
            Next sbr
            con.Execute cStrUpdate
                        
            con.Execute " DELETE FROM FILE1_10 WHERE ITEM = " & nBar2
            MsgBox " „ Õ–› «·»«—ﬂÊœ «·„ﬂ——"
        End If
    End If
End With
End Sub

Private Sub CMD_FIX_Click()
Dim nBaR1 As Double
Dim nBar2 As Double
With grid2
    If grid2.Rows >= 3 Then
        If MsgBox("Ÿ»ÿ »«—ﬂÊœ „ﬂ——", vbYesNo + vbDefaultButton2) = vbYes Then
            MsgBox "ÌÃ» ÿ»«⁄… »«—ﬂÊœ ··«‰’«› «· Ï ·Â« »«—ﬂÊœ „ﬂ—— Œÿ√"
            nBaR1 = Val(.TextMatrix(1, 0))
            nBar2 = Val(.TextMatrix(2, 0))
            If nBaR1 > 0 And nBar2 > 0 Then
                con.Execute " UPDATE FILE6_20   SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
                con.Execute " UPDATE FILE7_10   SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
                con.Execute " UPDATE FILE7_20   SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
                con.Execute " UPDATE FILE1_60       SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
                con.Execute " UPDATE FILE0_10       SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
                con.Execute " UPDATE FILE0_100      SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
                
                con.Execute " UPDATE FILE6_50       SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
                con.Execute " UPDATE FILE6_51       SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
                con.Execute " UPDATE FILE6_52       SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
                con.Execute " UPDATE FILE6_55       SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
                
                con.Execute " UPDATE FILE1_11_ALL   SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
                
                con.Execute " UPDATE FR7_20         SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
                con.Execute " UPDATE FR7_10         SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
                con.Execute " UPDATE FR6_20         SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
                con.Execute " UPDATE FR0_10         SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
                con.Execute " UPDATE FR0_100        SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
                con.Execute " UPDATE FR1_60         SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
                
                cString = "INSERT INTO BARCODE_EDIT (ITEMNEW ,ITEMOLD ) VALUES (  " & nBaR1 & " , " & nBar2 & " ) "
                con.Execute cString
                
                cStrUpdate = ""
                For sbr = 3 To nCountBranch
                    cBr = RetZero(sbr, 2)
                    cStrUpdate = cStrUpdate & " INSERT INTO NEW_ITEM (ITEM ,  BRANCH) VALUES        ( " & addvalue(nBaR1) & " , " & addstring(cBr) & " )  ; "
                Next sbr
                con.Execute cStrUpdate
                            
                con.Execute " DELETE FROM FILE1_10 WHERE ITEM = " & nBar2
                Inform " „ Õ–› «·»«—ﬂÊœ «·„ﬂ——"
            End If
        End If
    End If
End With
End Sub

Private Sub cmdBarCode_Click()

con.BeginTrans
On Error GoTo myerror
If Not IsNumeric(grid2.TextMatrix(grid2.Row, 0)) Then Exit Sub
cString = "INSERT INTO ADDPRINT(ITEM,BRANCH,QUANT,ISPRINT)"
cString = cString & _
          " VALUES ( " & grid2.TextMatrix(grid2.Row, 0) & " , " & addstring(cBranch) & ",  1,1 ) "
con.Execute cString
con.CommitTrans
Inform " „ «· ÕÊÌ· »‰Ã«Õ"
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub
Private Sub Form_Load()
Dim cString As String
Frame1.Visible = bopt2
openCon con
myloadGrdModel " FILE1_10.modelno = " & MyParn(pModelNo)
myload
Set grid2.DataSource = data20
data20.ConnectionString = strCon
grid2.Rows = 1
Fixgrd2
End Sub
Private Function myloadGrdModel(pFilter As String) As Boolean
Dim aRet As Variant, cFieldas As String, cField As String, cFilter As String

mydefineGrdModel
aRet = retFields(pFilter)
If aRet(0) = "" Then Exit Function

cField = aRet(0)
cFieldas = aRet(1)

FillItem cFieldas, cField, pFilter
FixCost cFieldas, cField, pFilter
fixGrdModel
myloadgrd = True
End Function
Private Sub fixGrdModel()
With grdModel
.ColHidden(0) = True
.ColWidth(1) = 1300
nColWidth = (.Width - 200 - .ColWidth(0) - .ColWidth(1)) / .Cols
If nColWidth < 500 Then nColWidth = 500
If nColWidth > 1200 Then nColWidth = 1200
For nCol = 2 To .Cols - 1
    .ColWidth(nCol) = nColWidth
    .ColAlignment(nCol) = flexAlignCenterCenter
Next
.RowHidden(1) = (cBranch <> "00")
'.RowHidden(2) = True
End With
End Sub
Private Sub mydefineGrdModel()
grdModel.Rows = 4
grdModel.Cols = 2

grdModel.MergeCells = flexMergeRestrictRows
grdModel.TextMatrix(0, 1) = "«·„ﬁ«”"
grdModel.TextMatrix(1, 1) = "”⁄— „’‰⁄"
grdModel.TextMatrix(2, 1) = "”⁄— „” Â·ﬂ"
grdModel.TextMatrix(3, 1) = "”⁄— «Êﬂ«“ÌÊ‰"
grdModel.RowHidden(1) = (cBranch <> "00")
End Sub
Private Sub FixCost(cFieldas, cField, pFilter As String)
Dim loctable As ADODB.Recordset
' „·∆ «·ÃœÊ·
cString = "Select " & cFieldas & _
          " From " & _
          " (Select scal,cost from file1_10 " & turn(pFilter) & pFilter & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(cost)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  "

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    For nCol = 2 To grdModel.Cols - 2
        grdModel.TextMatrix(1, nCol) = loctable.Fields(nCol - 2).Value & ""
    Next
End If

' „·∆ «·ÃœÊ·
cString = "Select " & cFieldas & _
          " From " & _
          " (Select scal,price from file1_10 " & turn(pFilter) & pFilter & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(price)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  "

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    For nCol = 2 To grdModel.Cols - 2
        grdModel.TextMatrix(2, nCol) = loctable.Fields(nCol - 2).Value & ""
    Next
End If
Set loctable = Nothing

' „·∆ «·ÃœÊ·
cString = "Select " & cFieldas & _
          " From " & _
          " (Select scal,price2 from file1_10 " & turn(pFilter) & pFilter & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(price2)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  "

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    For nCol = 2 To grdModel.Cols - 2
        grdModel.TextMatrix(3, nCol) = loctable.Fields(nCol - 2).Value & ""
    Next
End If
Set loctable = Nothing
End Sub
Private Function retFields(pFilter As String)
Dim aRet(1) As String
Dim FieldTable As New ADODB.Recordset
'  ⁄—Ì› «·«⁄„œ…
Dim cString As String, cWhere As String
FieldTable.Open "Select SCAL , MIN(C_SCAL) AS C_SCAL from file1_10 " & turn(pFilter) & pFilter & " GROUP BY SCAL order by MIN(c_scal) ", con, adOpenStatic, adLockReadOnly
Do Until FieldTable.EOF
    If Not IsNull(FieldTable!scal) Then
        cFieldas = cFieldas & turn(cField, ",") & "[" & FieldTable!scal & "]" & " as " & "[" & FieldTable!scal & "]"
        cField = cField & turn(cField, ",") & "[" & FieldTable!scal & "]"
    End If
    FieldTable.MoveNext
Loop

aRet(0) = cField
aRet(1) = cFieldas
retFields = aRet
' ⁄œ„ ÊÃÊœ «⁄„œ…
FieldTable.Close
Set FieldTable = Nothing
End Function
Private Sub FillItem(cFieldas, cField, pFilter As String)
Dim GrdTable As New ADODB.Recordset, cString As String
' „·∆ «·ÃœÊ·
cString = "Select c_color as [—ﬁ„ «··Ê‰] ,color as [«··Ê‰] " & turn(cFieldas, ",") & cFieldas & _
          " From " & _
          " (Select c_color,Color,(scal),item from file1_10 " & turn(pFilter) & pFilter & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(item)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  " & _
          " order by pvt.C_color"

GrdTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
grdModel.Cols = GrdTable.Fields.Count

For nCol = 2 To GrdTable.Fields.Count - 1
    grdModel.TextMatrix(0, nCol) = GrdTable.Fields(nCol).Name & ""
Next

Do Until GrdTable.EOF
    grdModel.AddItem ""
    For nCol = 0 To 1
        grdModel.TextMatrix(grdModel.Rows - 1, nCol) = GrdTable.Fields(nCol).Value & ""
    Next
    GrdTable.MoveNext
Loop

With grdModel
.AddItem ""
.TextMatrix(.Rows - 1, 1) = "≈Ã„«·Ì „ﬁ«”"
.Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HE0E0E0
.ColWidth(.Cols - 1) = 1000

.Cols = .Cols + 1
.TextMatrix(0, .Cols - 1) = "≈Ã„«·Ì ·Ê‰"
.TextMatrix(1, .Cols - 1) = "≈Ã„«·Ì ·Ê‰"
.TextMatrix(2, .Cols - 1) = "≈Ã„«·Ì ·Ê‰"
.TextMatrix(3, .Cols - 1) = "≈Ã„«·Ì ·Ê‰"
.MergeCells = flexMergeFixedOnly
.MergeCol(.Cols - 1) = True
.Cell(flexcpBackColor, 0, .Cols - 1, .Rows - 1, .Cols - 1) = &HE0E0E0
End With
GrdTable.Close
Set GrdTable = Nothing
End Sub
Private Sub GRID1_Click()
myload
End Sub
Private Sub CalcTotals()
With grdModel
If .Rows > 4 Then
    For nCol = 2 To .Cols - 1
        For nRow = 4 To .Rows - 2
            nTotal = nTotal + Val(.TextMatrix(nRow, nCol))
        Next
        .TextMatrix(.Rows - 1, nCol) = Myvalue(nTotal)
        nTotal = 0
    Next
    For nRow = 4 To .Rows - 1
        For nCol = 2 To .Cols - 2
            nTotal = nTotal + Val(.TextMatrix(nRow, nCol))
        Next
        .TextMatrix(nRow, .Cols - 1) = Myvalue(nTotal)
        nTotal = 0
    Next
    .ColAlignment(.Cols - 1) = flexAlignRightCenter
    .RowHidden(.Rows - 1) = True
End If
End With
End Sub
Private Sub myload()
For nRow = 4 To grdModel.Rows - 1
    For nCol = 2 To grdModel.Cols - 1
        grdModel.TextMatrix(nRow, nCol) = ""
    Next
Next

Dim loctable As New ADODB.Recordset, cString As String
loctable.Open pString, con, adOpenStatic, adLockReadOnly
Do Until loctable.EOF
    For nRow = 4 To grdModel.Rows - 1
        If grdModel.TextMatrix(nRow, 1) = loctable!Color Then
            For nCol = 2 To grdModel.Cols - 1
                If grdModel.TextMatrix(0, nCol) = loctable!scal Then
                    grdModel.TextMatrix(nRow, nCol) = loctable(pFieldValue) & ""
                    Exit For
                End If
            Next
        End If
    Next
    loctable.MoveNext
Loop
CalcTotals
loctable.Close
Set loctable = Nothing
End Sub
Private Sub myLoad_incase()
'For nRow = 2 To grdModelCodes.Rows - 1
'    For nCol = 2 To grdModelCodes.Cols - 1
'        grdModel.TextMatrix(nRow, nCol) = ""
'    Next
'Next
'
'If Val(grid1.TextMatrix(grid1.Row, 3)) = 0 And Val(grid1.TextMatrix(grid1.Row, 4)) = 0 Then Exit Sub
'Dim loctable As New ADODB.Recordset, cString As String
'If Val(grid1.TextMatrix(grid1.Row, 3)) <> 0 Then
'    cString = "SELECT SUM(QUANT) AS sumOfQuant,FILE1_10.COLOR,FILE1_10.SCAL FROM FILE7_20 INNER JOIN FILE1_10 ON FILE7_20.ITEM = FILE1_10.ITEM"
'    cString = cString & turn(cString) & "FILE7_20.DOC_NO = " & MyParn(grid1.TextMatrix(grid1.Row, 1))
'    cString = cString & turn(cString) & " FILE1_10.modelno = " & MyParn(VsTItem.xmodelno.Caption)
'Else
'    cString = "SELECT SUM(QUANT) as sumofQuant,FILE1_10.COLOR,FILE1_10.SCAL FROM FILE7_10 INNER JOIN FILE1_10 ON FILE7_10.ITEM = FILE1_10.ITEM"
'    cString = cString & turn(cString) & "FILE7_10.DOC_NO = " & MyParn(grid1.TextMatrix(grid1.Row, 1))
'    cString = cString & turn(cString) & " FILE1_10.modelno = " & MyParn(VsTItem.xmodelno.Caption)
'End If
'cString = cString & " GROUP BY FILE1_10.COLOR,FILE1_10.SCAL "
'loctable.Open cString, con, adOpenStatic, adLockReadOnly
'Do Until loctable.EOF
'    For nRow = 3 To grdModel.Rows - 1
'        If grdModel.TextMatrix(nRow, 1) = loctable!Color Then
'            For nCol = 2 To grdModel.Cols - 1
'                If grdModel.TextMatrix(0, nCol) = loctable!scal Then
'                    grdModel.TextMatrix(nRow, nCol) = loctable!sumofQuant
'                    Exit For
'                End If
'            Next
'        End If
'    Next
'    loctable.MoveNext
'Loop
'CalcTotals
'loctable.Close
'Set loctable = Nothing
End Sub
Private Sub FillItem_INCASE(cFieldas, cField, pFilter As String)
'Dim GRDTABLE As New ADODB.Recordset, cString As String
'' „·∆ «·ÃœÊ·
'cString = "Select c_color as [—ﬁ„ «··Ê‰] ,color as [«··Ê‰] " & turn(cFieldas, ",") & cFieldas & _
'          " From " & _
'          " (Select c_color,Color,scal,item from file1_10 " & turn(pFilter) & pFilter & " ) AS TABLE1" & _
'          " PIVOT " & _
'          " (max(item)" & _
'          " FOR SCAL IN " & _
'          "(" & cField & ")" & _
'          ") as pvt  " & _
'          " order by pvt.C_color"
'
'GRDTABLE.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
'grdModel.Cols = GRDTABLE.Fields.Count: grdModelCodes.Cols = GRDTABLE.Fields.Count
'
'For nCol = 2 To GRDTABLE.Fields.Count - 1
'    grdModel.TextMatrix(0, nCol) = GRDTABLE.Fields(nCol).Name
'Next
'
'Do Until GRDTABLE.EOF
'    grdModelCodes.AddItem ""
'    grdModel.AddItem ""
'    For nCol = 0 To GRDTABLE.Fields.Count - 1
'        If nCol <= 1 Then
'            grdModel.TextMatrix(grdModelCodes.Rows - 1, nCol) = GRDTABLE.Fields(nCol).Value & ""
'        Else
'            grdModelCodes.TextMatrix(grdModelCodes.Rows - 1, nCol) = GRDTABLE.Fields(nCol).Value & ""
'            grdModel.TextMatrix(grdModel.Rows - 1, nCol) = ""
'        End If
'    Next
'    GRDTABLE.MoveNext
'Loop
'With grdModel
'.AddItem ""
'.TextMatrix(.Rows - 1, 1) = "≈Ã„«·Ì „ﬁ«”"
'.Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HE0E0E0
'.ColWidth(.Cols - 1) = 1000
'
'.Cols = .Cols + 1
'.TextMatrix(0, .Cols - 1) = "≈Ã„«·Ì ·Ê‰"
'.Cell(flexcpBackColor, 0, .Cols - 1, .Rows - 1, .Cols - 1) = &HE0E0E0
'End With
'GRDTABLE.Close
'Set GRDTABLE = Nothing
End Sub
Sub MYLOAD2()
Dim cModelFact As String
Dim cFact As String
    '                   0                   1               2
'    cString = " SELECT file1_10.item , FILE4_10.desca ,  FILE0_40.DESCA , " & _
'            " sum([in] - out) as bal   " & _
'            " FROM   FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM INNER JOIN FILE4_10 ON FILE4_10.CODE = FILE1_10.CODE INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE1_11.STORE WHERE FILE1_10.MODELNO = " & MyParn(pModelNo) & _
'            " and scal = " & MyParn(grdModel.TextMatrix(0, grdModel.Col)) & _
'            " and color = " & MyParn(grdModel.TextMatrix(grdModel.Row, 1))
'    cString = cString & " GROUP BY FILE1_10.item, FILE4_10.DESCA ,  FILE0_40.DESCA"

    cModelFact = GetDesca("SELECT modelfact FROM FILE1_10 WHERE MODELNO = " & MyParn(pModelNo), con)
    cFact = GetDesca("SELECT fact FROM FILE1_10 WHERE MODELNO = " & MyParn(pModelNo), con)
    cString = " SELECT file1_10.item , FILE1_10.mosm  ,  ' ' , " & _
            " sum([in] - out) as bal   " & _
            " FROM   FILE1_10 LEFT JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM WHERE FILE1_10.modelfact = " & MyParn(cModelFact) & _
            " and fact = " & MyParn(cFact) & _
            " and scal = " & MyParn(grdModel.TextMatrix(0, grdModel.Col)) & _
            " and color = " & MyParn(grdModel.TextMatrix(grdModel.Row, 1))
    cString = cString & " GROUP BY FILE1_10.item, FILE1_10.mosm ORDER BY file1_10.item "
    data20.RecordSource = cString
    data20.Refresh
    Fixgrd2
End Sub
Sub Fixgrd2()
    With grid2
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .FixedCols = 0
    .Cols = 4
    .TextMatrix(0, 0) = "»«—ﬂÊœ"
    .TextMatrix(0, 1) = "„Ê”„"
    .TextMatrix(0, 2) = "«·›—⁄"
    .TextMatrix(0, 3) = "—’Ìœ"
    .ColWidth(0) = 1000
    .ColWidth(1) = 4000
    .ColWidth(2) = 2000
    .ColWidth(3) = 1000
    End With
End Sub
Private Sub grdModel_DblClick()
    With grdModel
        If .Row > 3 And .Row < .Rows - 1 Then
            MYLOAD2
        End If
    End With
End Sub
