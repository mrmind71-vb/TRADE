Attribute VB_Name = "lang"
Public Declare Function Beep Lib "kernel32" _
  (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
Declare Function GetLastError Lib "kernel32" () As Long

Const KLF_ACTIVATE = &H1

'Languages
Public Const Lang_AR As String = "00000401" 'Arabic
Public Const Lang_EN As String = "00000409" 'English

'Will return True if succeeds !
Public Function SetKbLayout(strLocaleId As String) As Boolean
    
    'Changes the KeyboardLayout
    'Returns TRUE when the KeyboardLayout was adjusted properly, FALSE otherwise
    'If the KeyboardLayout isn't installed, this function will install it for you
    On Error Resume Next
    Dim strLocId As String 'used to retrieve current KeyboardLayout

    'create a buffer
    strLocId = String(9, 0)
    'retrieve the current KeyboardLayout
    GetKeyboardLayoutName strLocId
    'Check whether the current KeyboardLayout and the
    'new one are the same
    If strLocId = (strLocaleId & Chr(0)) Then
        'If they're the same, we return immediately
        SetKbLayout = True
        Exit Function
    Else
        'create buffer
        strLocId = String(9, 0)
        
        'load and activate the layout for the current thread
        strLocId = LoadKeyboardLayout((strLocaleId & Chr(0)), KLF_ACTIVATE)

    End If
    
    'Test success

    GetKeyboardLayoutName strLocId
    
    If strLocId = (strLocaleId) Then SetKbLayout = True

End Function

