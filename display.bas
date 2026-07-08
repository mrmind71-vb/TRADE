Attribute VB_Name = "display2"
Private Declare Function SER_OPEN& Lib "SERSEN04.DLL" (ByVal commPort&, ByVal baudRate&)
Private Declare Function SER_CLOSE& Lib "SERSEN04.DLL" ()
Private Declare Sub SER_SEND Lib "SERSEN04.DLL" (theString$)
Private Declare Sub SER_RAWSEND Lib "SERSEN04.DLL" (theString$)
Private Declare Function SER_VERSINFO$ Lib "SERSEN04.DLL" ()
Sub display(Optional ByVal cText1 = "", Optional ByVal cText2 = "")
On Error Resume Next
Call SER_OPEN(1, 9600)      ' Comm 2, 9600 baud.
SER_SEND ("<CLEAR>")

If Not (Trim(cText1) = "" And Trim(cText2) = "") Then
    If Len(cText1) > 19 Then cText1 = Left(cText1, 19)
    If Len(cText2) > 19 Then cText2 = Left(cText2, 19)
    'SER_SEND ("<CLEAR>")
    cText = cText1 & IIf(cText1 <> "" And cText2 <> "", vbCrLf, "") & cText2
    SER_SEND (cText)
    ' SER_SEND (Text2.Text)
     'If Trim(Text3.Text) <> "" Then SER_SEND ("<" & Trim(Text3.Text) & ">")
    ' SER_SEND ("<LINE>" + Chr$(A + 64) + Chr$(B + 64) + Chr$(C + 64) + Chr$(D + 64))
End If
bClosed = SER_CLOSE              ' Close the port.
Exit Sub
myerror:
  MsgBox Err.Description
  Err.Clear
End Sub

