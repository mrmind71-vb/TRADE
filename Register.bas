Attribute VB_Name = "Protection"
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Const NCODE1 = 74, NCODE2 = 71
Public isDemo As Boolean
Function Serial(sDrive As String) As String
Dim tmp$
Dim j
Dim lpVolumeNameBuffer As String
Dim lpVolumeSerialNumber As Long
Dim lpMaximumComponentLength As Long
Dim lpfileSystemFlag As Long
Dim lpFileSystemNameBuffer As String
Dim nFileSystemNameSize As Long
sDrive = Left$(sDrive, 1) + ":\"
lpVolumeNameBuffer = Space$(128)
GetVolumeInformation sDrive, lpVolumeNameBuffer, _
                      Len(lpVolumeNameBuffer), lpVolumeSerialNumber, _
                      lpMaximumComponentLength, lpFileSystemFlags, _
                      lFileSystemNameBuffer, nFileSystemNameSize
Serial = Abs(lpVolumeSerialNumber)
End Function
Function UnCodeSerial(pSerial, pKey) As Double
firstnumber = Mid(pSerial, 1, 1)
UnCodeSerial = Mid(pSerial, 2)
UnCodeSerial = UnCodeSerial - (firstnumber * pKey)
UnCodeSerial = UnCodeSerial / firstnumber
End Function
Function CodeSerial(pSerial, pKey) As Double
Dim firstnumber
Randomize
firstnumber = Int((9 * Rnd) + 1)     ' Generate random value between 1 and 6.
CodeSerial = firstnumber * pSerial
CodeSerial = CodeSerial + (firstnumber * pKey)
CodeSerial = firstnumber & CodeSerial
End Function
Function isregistered() As Boolean
On Error GoTo myerror
Open App.Path & "\serial.txt" For Input As #1
Input #1, mystring
Close #1
If Serial("C") = UnCodeSerial(mystring, NCODE2) Then
    isregistered = True
Else
    GoTo myerror
End If
Exit Function
myerror:
    MsgBox "«·‰”Œ… €Ì— „—Œ’…", , systemName
    isregistered = False
End Function


