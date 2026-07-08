Attribute VB_Name = "ApiDeclarations"
Option Explicit
Dim OSSTab1 As Integer
Dim IntrCh23$, IntrCh24$, IntrCh40$, IntrCh5B$, IntrCh5C$, IntrCh5D$, IntrCh5E$
Dim IntrCh60$, IntrCh7B$, IntrCh7C$, IntrCh7D$, IntrCh7E$

' **********************************************
' Posiflex usbpd.dll DLL
' **********************************************
Public Declare Function WritePD _
    Lib "usbpd.dll" _
    (ByVal data As String, ByVal length As Long) _
As Long

Public Declare Function WritePD80 _
    Lib "usbpd.dll" Alias "WritePD" _
    (ByRef data As Any, ByVal length As Long) _
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

Declare Sub Sleep Lib "kernel32" _
   (ByVal dwMilliseconds As Long)

