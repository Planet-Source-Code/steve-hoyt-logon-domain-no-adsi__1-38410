Attribute VB_Name = "basCurrentDomain"
Option Explicit

Private Const VER_PLATFORM_WIN32_NT                 As Long = 2
Private Const HKEY_LOCAL_MACHINE                    As Long = &H80000002
Private Const KEY_READ                              As Long = &H20019
Private Const mcstrAgentKey                         As String * 56 = "System\CurrentControlSet\Services\MSNP32\NetworkProvider"

Private Type OS_VERSION_INFO
    dwOSVersionInfoSize     As Long
    dwMajorVersion          As Long
    dwMinorVersion          As Long
    dwBuildNumber           As Long
    dwPlatformId            As Long
    szCSDVersion            As String * 128
End Type

Private Type WKSTA_USER_INFO_1
    wkui1_username          As Long
    wkui1_logon_domain      As Long
    wkui1_oth_domains       As Long
    wkui1_logon_server      As Long
End Type

'Common APIs
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetPlatform Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OS_VERSION_INFO) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'WinNT/2000/XP APIs
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function NetApiBufferFree Lib "Netapi32.dll" (ByVal lpBuffer As Long) As Long
Private Declare Function NetWkstaUserGetInfo Lib "Netapi32.dll" (ByVal reserved As Any, ByVal Level As Long, lpBuffer As Any) As Long

'Win9x/ME APIs
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Private Sub Main()
    MsgBox CurrentDomain
End Sub

Public Function CurrentDomain() As String
    Dim lngStructPtr        As Long
    Dim udtUserInfo         As WKSTA_USER_INFO_1
    Select Case NetApiSupport
        Case True:  NetWkstaUserGetInfo 0&, 1&, lngStructPtr
        Case False: Win9xDomainName CurrentDomain
    End Select
    If lngStructPtr = 0 Then Exit Function
    CopyMemory udtUserInfo, ByVal lngStructPtr, Len(udtUserInfo)
    CurrentDomain = StrFromPtr(udtUserInfo.wkui1_logon_domain)
    NetApiBufferFree lngStructPtr
End Function

Private Function StrFromPtr(lngPtr As Long) As String
    Dim bytString()      As Byte
    Dim lngBytes         As Long
    If lngPtr = 0 Then Exit Function
    lngBytes = lstrlenW(lngPtr) * 2
    If lngBytes = 0 Then Exit Function
    ReDim bytString(0 To (lngBytes - 1)) As Byte
    CopyMemory bytString(0), ByVal lngPtr, lngBytes
    StrFromPtr = bytString
End Function

Private Function NetApiSupport() As Boolean
    On Error Resume Next
    Dim udtOS       As OS_VERSION_INFO
    udtOS.dwOSVersionInfoSize = Len(udtOS)
    If Not (GetPlatform(udtOS) = 1) Then Exit Function
    NetApiSupport = (udtOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Private Sub Win9xDomainName(ByRef strDomainName As String)
    On Error Resume Next
    Dim strRegKey           As String
    Dim lngRegKeyHdl        As Long
    Dim lngRegKeyLen        As Long
    Dim lngRegDataType      As Long
    strRegKey = Space(255)
    lngRegKeyLen = Len(strRegKey)
    RegOpenKeyEx HKEY_LOCAL_MACHINE, ByVal mcstrAgentKey, 0&, KEY_READ, lngRegKeyHdl
    RegQueryValueEx lngRegKeyHdl, "AuthenticatingAgent", 0, lngRegDataType, ByVal strRegKey, lngRegKeyLen
    RegCloseKey lngRegKeyHdl
    strDomainName = Left$(strRegKey, lngRegKeyLen - 1)
    strDomainName = Trim$(strDomainName)
End Sub

