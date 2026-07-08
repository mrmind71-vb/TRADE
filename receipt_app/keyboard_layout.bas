Attribute VB_Name = "mdKdeyBoardLayUt"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2011 VBnet/Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const LOCALE_ILANGUAGE             As Long = &H1     'language id
Private Const LOCALE_SLANGUAGE             As Long = &H2     'localized name of language
Private Const LOCALE_SENGLANGUAGE          As Long = &H1001  'English name of language
Private Const LOCALE_SABBREVLANGNAME       As Long = &H3     'abbreviated language name
Private Const LOCALE_SCOUNTRY              As Long = &H6     'localized name of country
Private Const LOCALE_SENGCOUNTRY           As Long = &H1002  'English name of country
Private Const LOCALE_SABBREVCTRYNAME       As Long = &H7     'abbreviated country name
'#if(WINVER >=  &H0400)
Private Const LOCALE_SISO639LANGNAME       As Long = &H59    'ISO abbreviated language name
Private Const LOCALE_SISO3166CTRYNAME      As Long = &H5A    'ISO abbreviated country name

Public Enum kbMode
    ILANGUAGE = 1
    SLANGUAGE = 2
    SENGLANGUAGE = 3
    SABBREVLANGNAME = 4
    SCOUNTRY = 5
    SENGCOUNTRY = 6
    SABBREVCTRYNAME = 7
    SISO3166CTRYNAME = 8
    SISO639LANGNAME = 9
End Enum
Private Declare Function GetKeyboardLayout Lib "user32" _
  (ByVal dwLayout As Long) As Long

Private Declare Function GetLocaleInfo Lib "kernel32" _
   Alias "GetLocaleInfoA" _
  (ByVal Locale As Long, _
   ByVal LCType As Long, _
   ByVal lpLCData As String, _
   ByVal cchData As Long) As Long
Public Function GetKeyBoard(pMode As kbMode)
Dim hKeyboardID As Long
Dim LCID As Long

'Identifies the thread to query, or is
'zero for the current thread.
hKeyboardID = GetKeyboardLayout(0&)

If hKeyboardID > 0 Then

    LCID = LoWord(hKeyboardID)

    If LCID Then
      Select Case pMode
          Case kbMode.ILANGUAGE
              GetKeyBoard = GetUserLocaleInfo(LCID, LOCALE_ILANGUAGE)
          Case kbMode.SLANGUAGE
              GetKeyBoard = GetUserLocaleInfo(LCID, LOCALE_SLANGUAGE)
          Case kbMode.SENGLANGUAGE
              GetKeyBoard = GetUserLocaleInfo(LCID, LOCALE_SENGLANGUAGE)
          Case kbMode.SABBREVLANGNAME
              GetKeyBoard = GetUserLocaleInfo(LCID, LOCALE_SABBREVLANGNAME)
          Case kbMode.SCOUNTRY
              GetKeyBoard = GetUserLocaleInfo(LCID, LOCALE_SCOUNTRY)
          Case kbMode.SENGCOUNTRY
              GetKeyBoard = GetUserLocaleInfo(LCID, LOCALE_SENGCOUNTRY)
          Case kbMode.SABBREVCTRYNAME
              GetKeyBoard = GetUserLocaleInfo(LCID, LOCALE_SABBREVCTRYNAME)
          Case kbMode.SISO3166CTRYNAME
              GetKeyBoard = GetUserLocaleInfo(LCID, LOCALE_SISO3166CTRYNAME)
          Case kbMode.SISO639LANGNAME
              GetKeyBoard = GetUserLocaleInfo(LCID, LOCALE_SISO639LANGNAME)
      End Select
    End If
End If
End Function


Private Function LoWord(wParam As Long) As Integer

   If wParam And &H8000& Then
      LoWord = &H8000& Or (wParam And &H7FFF&)
   Else
      LoWord = wParam And &HFFFF&
   End If

End Function


Public Function GetUserLocaleInfo(ByVal dwLocaleID As Long, _
                                  ByVal dwLCType As Long) As String

   Dim sReturn As String
   Dim nSize As Long

  'call the function passing the Locale type
  'variable to first retrieve the required
  'size of the string buffer needed
   nSize = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    
  'if successful (nSize > 0)
   If nSize > 0 Then
    
     'pad a buffer with spaces
      sReturn = Space$(nSize)
       
     'and call again passing the buffer
      nSize = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
     
     'if successful (nSize > 0)
      If nSize > 0 Then
      
        'nSize holds the size of the string
        'including the terminating null
         GetUserLocaleInfo = Left$(sReturn, nSize - 1)
      
      End If
   
   End If
    
End Function

