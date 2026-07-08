Attribute VB_Name = "mShellExecute"
Option Explicit
Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Const SEE_MASK_NOCLOSEPROCESS = &H40

Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1
Public Const SW_MINIMIZE = 2
Public Const SW_MAXIMIZE = 3

Public Const SE_ERR_FNF = 2
Public Const SE_ERR_NOASSOC = 31
Public Const SE_ERR_ACCESSDENIED = 5
Public Const INFINITE = &HFFFF
Public Const WAIT_TIMEOUT = &H102

Public Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO) As Long
Public Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long


Public Function FileExists(ByVal sFileName As String) As Boolean
    Dim intReturn As Integer

    On Error GoTo FileExists_Error
    intReturn = GetAttr(sFileName)
    FileExists = True
    
    Exit Function

FileExists_Error:
    FileExists = False
End Function
Public Sub Main()
End Sub

Public Function PathOfFile(FileName As String) As String
    Dim posn As Integer
    
    posn = InStrRev(FileName, "\")
    If posn > 0 Then
        PathOfFile = Left$(FileName, posn)
    Else
        PathOfFile = ""
    End If

End Function
Public Function ShellExWait(File As String, Parameter As String, pform As Form, Optional bVisible As Boolean = False) As Boolean
    Dim sei As SHELLEXECUTEINFO
    Dim retVal As Long
    
    If Not FileExists(File) Then
        ShellExWait = False
        Exit Function
    End If
    
    With sei
        .cbSize = Len(sei)
        .fMask = SEE_MASK_NOCLOSEPROCESS
        .hwnd = pform.hwnd
        .lpVerb = "open"
        .lpFile = File
        .lpParameters = Parameter
        .lpDirectory = PathOfFile(File)
        .nShow = IIf(bVisible, SW_NORMAL, 0)
        '.nShow = IIf(bVisible, 2, 0)

    End With
    
    retVal = ShellExecuteEx(sei)
    If retVal = 0 Then
        Select Case sei.hInstApp
             Case SE_ERR_FNF
                 Debug.Print "The file was not found."
             Case SE_ERR_NOASSOC
                 Debug.Print "No program is associated with this kind of file."
             Case SE_ERR_ACCESSDENIED
                 Debug.Print "Access denied."
             Case Else
                 Debug.Print "An unexpected error occured."
            End Select
        ShellExWait = False
    Else
        Do
            DoEvents
            retVal = WaitForSingleObject(sei.hProcess, 0)
        Loop While retVal = WAIT_TIMEOUT
        ShellExWait = True
    End If
End Function

