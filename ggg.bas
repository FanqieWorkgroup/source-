Attribute VB_Name = "Module1"
'Ä£¿é´úÂë
Option Explicit
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Dim oDB As Object
Dim odbRS As Object
Declare Function FindFirstFile Lib "kernel32" Alias _
"FindFirstFileA" (ByVal lpFileName As String, lpFindFileData _
As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
(ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function GetFileAttributes Lib "kernel32" Alias _
"GetFileAttributesA" (ByVal lpFileName As String) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) _
As Long
Declare Function FileTimeToLocalFileTime Lib "kernel32" _
(lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Declare Function FileTimeToSystemTime Lib "kernel32" _
(lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Const MAX_PATH = 260
Public Const MAXDWORD = &HFFFF
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type
Type WIN32_FIND_DATA
dwFileAttributes As Long
ftCreationTime As FILETIME
ftLastAccessTime As FILETIME
ftLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
cFileName As String * MAX_PATH
cAlternate As String * 14
End Type
Type SYSTEMTIME
wYear As Integer
wMonth As Integer
wDayOfWeek As Integer
wDay As Integer
wHour As Integer
wMinute As Integer
wSecond As Integer
wMilliseconds As Integer
End Type
Public Function StripNulls(OriginalStr As String) As String
If (InStr(OriginalStr, Chr(0)) > 0) Then
OriginalStr = Left(OriginalStr, _
InStr(OriginalStr, Chr(0)) - 1)
End If
StripNulls = OriginalStr
End Function
