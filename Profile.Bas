Attribute VB_Name = "Profile_bas"
Option Explicit

    Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpDefault As String, ByVal lpReturnedString As String, _
         ByVal nSize As Long, ByVal lpFileName As String) As Long
    Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Function LeerIni(lpFileName As String, lpAppName As String, lpKeyName As String, Optional vDefault) As String
    'Read INI
    'Parameters:
    'lpFileName:    INI path
    'lpAppName:     Section
    'lpKeyName:     Key
    'vDefault:      default value
    '
    Dim lpString As String
    Dim LTmp As Long
    Dim sRetVal As String
    
    If IsMissing(vDefault) Then
        lpString = ""
    Else
        lpString = vDefault
    End If
    
    sRetVal = String$(255, 0)
    
    LTmp = GetPrivateProfileString(lpAppName, lpKeyName, lpString, sRetVal, Len(sRetVal), lpFileName)
    If LTmp = 0 Then
        LeerIni = lpString
    Else
        LeerIni = Left(sRetVal, LTmp)
    End If
End Function


Sub GuardarIni(lpFileName As String, lpAppName As String, lpKeyName As String, lpString As String)
    'Save INI
    Dim LTmp As Long

    LTmp = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
End Sub

