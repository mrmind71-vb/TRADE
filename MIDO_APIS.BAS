Attribute VB_Name = "ApiDeclarations"
' **********************************************
' Posiflex usbpd.dll DLL
' **********************************************
Public Declare Function WritePD _
    Lib "usbpd.dll" _
    (ByVal data As String, ByVal length As Long) _
As Long

Public Declare Function PdState _
    Lib "usbpd.dll" _
    () _
As Long

Public Declare Function OpenUSBpd _
    Lib "usbpd.dll" _
    () _
As Long

Public Declare Function CloseUSBpd _
    Lib "usbpd.dll" _
    () _
As Long
Function myDisplay(ByVal cString1 As String, Optional ByVal cString2 As String)
On Error GoTo myerror
Dim nError As Long
nError = OpenUSBpd()

Dim cString As String
cString = Chr(&HC) & String(40, Space(1))
nError = WritePD(cString, Len(cString))

If cString2 <> "" Then
    cString = cString1
    cString1 = Left(cString1, 20)
    cString2 = Left(cString2, 20)
    cString = cString1 & String(20 - Len(cString1), Space(1))
    cString = cString & cString2
Else
    cString = Left(cString1, 40)
End If
cString = Chr(&HC) & cString
nError = WritePD(cString, Len(cString))

DoEvents
nError = CloseUSBpd()
Exit Function
myerror:
Err.Clear
End Function

