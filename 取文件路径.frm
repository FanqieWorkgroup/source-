VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form form1 
   BackColor       =   &H8000000C&
   Caption         =   "����U�̱������"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8265
   LinkTopic       =   "����U�̱������"
   ScaleHeight     =   6600
   ScaleWidth      =   8265
   StartUpPosition =   3  '����ȱʡ
   Begin SysInfoLib.SysInfo SysInfo2 
      Left            =   240
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox SysInfo1 
      Height          =   480
      Left            =   240
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.ListBox List2 
      Height          =   2220
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   2580
      Left            =   6360
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000003&
      Caption         =   "����"
      Height          =   2295
      Left            =   1080
      TabIndex        =   3
      Top             =   3840
      Width           =   6735
      Begin VB.CommandButton Command4 
         Caption         =   "��ϵ����"
         Height          =   975
         Left            =   3720
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         Caption         =   "�������"
         Height          =   1095
         Left            =   480
         TabIndex        =   7
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "U�̱���"
      Height          =   3855
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "������������"
         Height          =   1575
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "������ɨ��"
         Height          =   1575
         Left            =   840
         TabIndex        =   2
         Top             =   2160
         Width           =   3495
      End
   End
   Begin VB.ListBox List3 
      Height          =   1500
      Left            =   360
      TabIndex        =   0
      Top             =   6840
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function DeleteFile Lib "KERNEL32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function GetDriveType Lib "KERNEL32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Dim oDB As Object
Dim odbRS As Object
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function cloudr Lib "cloudsearch.dll" Alias "getresult" (ByVal md5_string As String) As Integer






Private Sub Command1_Click()
frmOptions.Show
End Sub

Private Sub Command2_Click()
Dialog.Show
Dim SearchPath As String, FindStr As String
Dim FileSize As Long
Dim NumFiles As Integer, NumDirs As Integer
Dim i As Integer
Screen.MousePointer = vbHourglass
List3.Clear
Call USBDISK
For i = 0 To List3.ListCount - 1
SearchPath = List3.List(i) & ":\"
FindStr = "*.*"
If InternetGetConnectedState(0&, 0&) Then
FileSize = FindFilesAPIcloud2(SearchPath, FindStr, NumFiles, NumDirs)
Else
FindFilesAPI2 SearchPath, FindStr, NumFiles, NumDirs
End If
Screen.MousePointer = vbDefault
Next i
End Sub


 Function USBDISK()
Dim i As Long
    For i = Asc("A") To Asc("Z")
        If GetDriveType(Chr(i) + ":") = 2 Then List3.AddItem Chr(i)
    Next i
End Function



'����microsoft scprting runtime

'�������

Function FindFilesAPIq(path As String, SearchStr As String, _
FileCount As Integer, DirCount As Integer) '��Ĭģʽ
Dim FileName As String ' Walking filename variable...
Dim DirName As String ' SubDirectory Name
Dim dirNames() As String ' Buffer for directory name entries
Dim nDir As Integer ' Number of directories in this path
Dim i As Integer ' For-loop counter...
Dim hSearch As Long ' Search Handle
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer
Dim FT As FILETIME
Dim ST As SYSTEMTIME
Dim DateCStr As String, DateMStr As String
If Right(path, 1) <> "\" Then path = path & "\"
' Search for subdirectories.
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(path & "*", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
Do While Cont
DirName = StripNulls(WFD.cFileName)
' Ignore the current and encompassing directories.
If (DirName <> ".") And (DirName <> "..") Then
' Check for directory with bitwise comparison.
If GetFileAttributes(path & DirName) And _
FILE_ATTRIBUTE_DIRECTORY Then
dirNames(nDir) = DirName
DirCount = DirCount + 1
nDir = nDir + 1
ReDim Preserve dirNames(nDir)
' Uncomment the next line to list directories
'List1.AddItem path & FileName
End If
End If
Cont = FindNextFile(hSearch, WFD) ' Get next subdirectory.
Loop
Cont = FindClose(hSearch)
End If
' Walk through this directory and sum file sizes.
hSearch = FindFirstFile(path & SearchStr, WFD)
Cont = True
If hSearch <> INVALID_HANDLE_VALUE Then
While Cont
FileName = StripNulls(WFD.cFileName)
If (FileName <> ".") And (FileName <> "..") And _
((GetFileAttributes(path & FileName) And _
FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
FindFilesAPIq = FindFilesAPIq + (WFD.nFileSizeHigh * _
MAXDWORD) + WFD.nFileSizeLow
FileCount = FileCount + 1
' To list files w/o dates, uncomment the next line
' and remove or Comment the lines down to End If
'List1.AddItem path & FileName
' Include Creation date...
FileTimeToLocalFileTime WFD.ftCreationTime, FT
FileTimeToSystemTime FT, ST
DateCStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
" " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
' and Last Modified Date
FileTimeToLocalFileTime WFD.ftLastWriteTime, FT
FileTimeToSystemTime FT, ST
DateMStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
" " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
List1.AddItem path & FileName & vbTab & _
Format(DateCStr, "mm/dd/yyyy hh:nn:ss") _
& vbTab & Format(DateMStr, "mm/dd/yyyy hh:nn:ss")
MD5File path & FileName
DoEvents
Set oDB = CreateObject("litex.liteconnection")
Set odbRS = CreateObject("LiteX.LiteStatement")
odbRS.ActiveConnection = oDB
oDB.Open (App.path & "\virnuscenter.db")
odbRS.ActiveConnection = oDB
 odbRS.Prepare ("select * from md5 where words like '%" & GetMD5Text() & "%' ")
 odbRS.Step
List2.AddItem odbRS.RowCount
 If odbRS.RowCount <> 0 Then
DeleteFile path & FileName
 End If
End If
Cont = FindNextFile(hSearch, WFD) ' Get next file
Wend
Cont = FindClose(hSearch)
End If
' If there are sub-directories...
If nDir > 0 Then
' Recursively walk into them...
For i = 0 To nDir - 1
FindFilesAPIq = FindFilesAPIq + FindFilesAPIq(path & dirNames(i) _
& "\", SearchStr, FileCount, DirCount)
Next i
End If
End Function





Private Sub Command3_Click()
MsgBox "�������̱������ V1.0��ʽ��    ������汾��db.2018051jf", , "����"
End Sub

Private Sub Command4_Click()
MsgBox "bug������835078903@qq.com,������qqͬ��,������fanqie.gq,���µ�ַ��fanqiesupportpage.gq", , "��ϵ����"


End Sub

Private Sub Form_Load()
frmOptions.Check1.Value = 1
Register (App.path & "\sqlite3.dll")
  Register (App.path & "\cloudsearch.dll")
End Sub

Private Sub SysInfo1_DeviceArrival(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
If frmOptions.Check1.Value = 1 Then
List3.Clear
 Dim SearchPath As String, FindStr As String
Dim FileSize As Long
Dim NumFiles As Integer, NumDirs As Integer
Dim i As Integer
Screen.MousePointer = vbHourglass

Call USBDISK
For i = 0 To List3.ListCount - 1
SearchPath = List3.List(i) & ":\"
FindStr = "*.*"
If InternetGetConnectedState(0&, 0&) Then
FileSize = FindFilesAPIcloud(SearchPath, FindStr, NumFiles, NumDirs)
Screen.MousePointer = vbDefault
Else
FindFilesAPI SearchPath, FindStr, NumFiles, NumDirs
End If
Next i
Else
End If
If frmOptions.Check1.Value = 1 And frmOptions.Check2.Value = 1 Then
Call USBDISK
For i = 0 To List3.ListCount - 1
SearchPath = List3.List(i) & ":\"
FindStr = "*.*"
If InternetGetConnectedState(0&, 0&) Then
FileSize = FindFilesAPIcloudq(SearchPath, FindStr, NumFiles, NumDirs)
Screen.MousePointer = vbDefault
Else
FindFilesAPIq SearchPath, FindStr, NumFiles, NumDirs
End If
Next i
Else
End If
End Sub

Function FindFilesAPI(path As String, SearchStr As String, _
FileCount As Integer, DirCount As Integer) 'if command2 click
Dim FileName As String ' Walking filename variable...
Dim DirName As String ' SubDirectory Name
Dim dirNames() As String ' Buffer for directory name entries
Dim nDir As Integer ' Number of directories in this path
Dim i As Integer ' For-loop counter...
Dim hSearch As Long ' Search Handle
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer
Dim FT As FILETIME
Dim ST As SYSTEMTIME
Dim DateCStr As String, DateMStr As String
If Right(path, 1) <> "\" Then path = path & "\"
' Search for subdirectories.
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(path & "*", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
Do While Cont
DirName = StripNulls(WFD.cFileName)
' Ignore the current and encompassing directories.
If (DirName <> ".") And (DirName <> "..") Then
' Check for directory with bitwise comparison.
If GetFileAttributes(path & DirName) And _
FILE_ATTRIBUTE_DIRECTORY Then
dirNames(nDir) = DirName
DirCount = DirCount + 1
nDir = nDir + 1
ReDim Preserve dirNames(nDir)
' Uncomment the next line to list directories
'List1.AddItem path & FileName
End If
End If
Cont = FindNextFile(hSearch, WFD) ' Get next subdirectory.
Loop
Cont = FindClose(hSearch)
End If
' Walk through this directory and sum file sizes.
hSearch = FindFirstFile(path & SearchStr, WFD)
Cont = True
If hSearch <> INVALID_HANDLE_VALUE Then
While Cont
FileName = StripNulls(WFD.cFileName)
If (FileName <> ".") And (FileName <> "..") And _
((GetFileAttributes(path & FileName) And _
FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * _
MAXDWORD) + WFD.nFileSizeLow
FileCount = FileCount + 1
' To list files w/o dates, uncomment the next line
' and remove or Comment the lines down to End If
'List1.AddItem path & FileName
' Include Creation date...
FileTimeToLocalFileTime WFD.ftCreationTime, FT
FileTimeToSystemTime FT, ST
DateCStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
" " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
' and Last Modified Date
FileTimeToLocalFileTime WFD.ftLastWriteTime, FT
FileTimeToSystemTime FT, ST
DateMStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
" " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
List1.AddItem path & FileName & vbTab & _
Format(DateCStr, "mm/dd/yyyy hh:nn:ss") _
& vbTab & Format(DateMStr, "mm/dd/yyyy hh:nn:ss")
MD5File path & FileName
DoEvents

Set oDB = CreateObject("litex.liteconnection")
Set odbRS = CreateObject("LiteX.LiteStatement")
odbRS.ActiveConnection = oDB
oDB.Open (App.path & "\virnuscenter.db")
odbRS.ActiveConnection = oDB
 odbRS.Prepare ("select * from md5 where words like '%" & GetMD5Text() & "%' ")
 odbRS.Step
List2.AddItem odbRS.RowCount
 If odbRS.RowCount <> 0 Then
 Form2.List1.AddItem path & FileName
 End If
End If
Cont = FindNextFile(hSearch, WFD) ' Get next file
Wend
Cont = FindClose(hSearch)
End If
' If there are sub-directories...
If nDir > 0 Then
' Recursively walk into them...
For i = 0 To nDir - 1
FindFilesAPI = FindFilesAPI + FindFilesAPI(path & dirNames(i) _
& "\", SearchStr, FileCount, DirCount)
Next i
End If
If List2.ListCount <> 0 Then
Form2.Show
End If
End Function
Function FindFilesAPI2(path As String, SearchStr As String, _
FileCount As Integer, DirCount As Integer)
Dialog.Label2.Caption = "����ʹ�ñ��ز�����"
Dim FileName As String ' Walking filename variable...
Dim DirName As String ' SubDirectory Name
Dim dirNames() As String ' Buffer for directory name entries
Dim nDir As Integer ' Number of directories in this path
Dim i As Integer ' For-loop counter...
Dim hSearch As Long ' Search Handle
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer
Dim FT As FILETIME
Dim ST As SYSTEMTIME
Dim DateCStr As String, DateMStr As String
If Right(path, 1) <> "\" Then path = path & "\"
' Search for subdirectories.
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(path & "*", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
Do While Cont
DirName = StripNulls(WFD.cFileName)
' Ignore the current and encompassing directories.
If (DirName <> ".") And (DirName <> "..") Then
' Check for directory with bitwise comparison.
If GetFileAttributes(path & DirName) And _
FILE_ATTRIBUTE_DIRECTORY Then
dirNames(nDir) = DirName
DirCount = DirCount + 1
nDir = nDir + 1
ReDim Preserve dirNames(nDir)
' Uncomment the next line to list directories
'List1.AddItem path & FileName
End If
End If
Cont = FindNextFile(hSearch, WFD) ' Get next subdirectory.
Loop
Cont = FindClose(hSearch)
End If
' Walk through this directory and sum file sizes.
hSearch = FindFirstFile(path & SearchStr, WFD)
Cont = True
If hSearch <> INVALID_HANDLE_VALUE Then
While Cont
FileName = StripNulls(WFD.cFileName)
If (FileName <> ".") And (FileName <> "..") And _
((GetFileAttributes(path & FileName) And _
FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
FindFilesAPI2 = FindFilesAPI2 + (WFD.nFileSizeHigh * _
MAXDWORD) + WFD.nFileSizeLow
FileCount = FileCount + 1
' To list files w/o dates, uncomment the next line
' and remove or Comment the lines down to End If
'List1.AddItem path & FileName
' Include Creation date...
FileTimeToLocalFileTime WFD.ftCreationTime, FT
FileTimeToSystemTime FT, ST
DateCStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
" " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
' and Last Modified Date
FileTimeToLocalFileTime WFD.ftLastWriteTime, FT
FileTimeToSystemTime FT, ST
DateMStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
" " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
List1.AddItem path & FileName & vbTab & _
Format(DateCStr, "mm/dd/yyyy hh:nn:ss") _
& vbTab & Format(DateMStr, "mm/dd/yyyy hh:nn:ss")
Dialog.Label1.Caption = "Scaning ��" & path & FileName
DoEvents
MD5File path & FileName
DoEvents
Set oDB = CreateObject("litex.liteconnection")
Set odbRS = CreateObject("LiteX.LiteStatement")
odbRS.ActiveConnection = oDB
oDB.Open (App.path & "\virnuscenter.db")
odbRS.ActiveConnection = oDB
 odbRS.Prepare ("select * from md5 where words like '%" & GetMD5Text() & "%' ")
 DoEvents
 odbRS.Step
List2.AddItem odbRS.RowCount
 If odbRS.RowCount <> 0 Then
 Form2.List1.AddItem path & FileName
 End If
End If
Cont = FindNextFile(hSearch, WFD) ' Get next file
Wend
Cont = FindClose(hSearch)
End If
' If there are sub-directories...
If nDir > 0 Then
' Recursively walk into them...
For i = 0 To nDir - 1
FindFilesAPI2 = FindFilesAPI2 + FindFilesAPI2(path & dirNames(i) _
& "\", SearchStr, FileCount, DirCount)
Next i
End If
If List2.ListCount <> 0 Then
Form2.Show
End If
End Function
Function FindFilesAPIcloud(path As String, SearchStr As String, _
FileCount As Integer, DirCount As Integer) 'if command2 click
Dim FileName As String ' Walking filename variable...
Dim DirName As String ' SubDirectory Name
Dim dirNames() As String ' Buffer for directory name entries
Dim nDir As Integer ' Number of directories in this path
Dim i As Integer ' For-loop counter...
Dim hSearch As Long ' Search Handle
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer
Dim FT As FILETIME
Dim ST As SYSTEMTIME
Dim DateCStr As String, DateMStr As String
If Right(path, 1) <> "\" Then path = path & "\"
' Search for subdirectories.
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(path & "*", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
Do While Cont
DirName = StripNulls(WFD.cFileName)
' Ignore the current and encompassing directories.
If (DirName <> ".") And (DirName <> "..") Then
' Check for directory with bitwise comparison.
If GetFileAttributes(path & DirName) And _
FILE_ATTRIBUTE_DIRECTORY Then
dirNames(nDir) = DirName
DirCount = DirCount + 1
nDir = nDir + 1
ReDim Preserve dirNames(nDir)
' Uncomment the next line to list directories
'List1.AddItem path & FileName
End If
End If
Cont = FindNextFile(hSearch, WFD) ' Get next subdirectory.
Loop
Cont = FindClose(hSearch)
End If
' Walk through this directory and sum file sizes.
hSearch = FindFirstFile(path & SearchStr, WFD)
Cont = True
If hSearch <> INVALID_HANDLE_VALUE Then
While Cont
FileName = StripNulls(WFD.cFileName)
If (FileName <> ".") And (FileName <> "..") And _
((GetFileAttributes(path & FileName) And _
FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
FindFilesAPIcloud = FindFilesAPIcloud + (WFD.nFileSizeHigh * _
MAXDWORD) + WFD.nFileSizeLow
FileCount = FileCount + 1
' To list files w/o dates, uncomment the next line
' and remove or Comment the lines down to End If
'List1.AddItem path & FileName
' Include Creation date...
FileTimeToLocalFileTime WFD.ftCreationTime, FT
FileTimeToSystemTime FT, ST
DateCStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
" " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
' and Last Modified Date
FileTimeToLocalFileTime WFD.ftLastWriteTime, FT
FileTimeToSystemTime FT, ST
DateMStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
" " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
List1.AddItem path & FileName & vbTab & _
Format(DateCStr, "mm/dd/yyyy hh:nn:ss") _
& vbTab & Format(DateMStr, "mm/dd/yyyy hh:nn:ss")
MD5File path & FileName
DoEvents
If cloudr(GetMD5Text()) = 3 Then
Form2.List1.AddItem path & FileName
End If
End If
Cont = FindNextFile(hSearch, WFD) ' Get next file
Wend
Cont = FindClose(hSearch)
End If
' If there are sub-directories...
If nDir > 0 Then
' Recursively walk into them...
For i = 0 To nDir - 1
FindFilesAPIcloud = FindFilesAPIcloud + FindFilesAPIcloud(path & dirNames(i) _
& "\", SearchStr, FileCount, DirCount)
Next i
End If
If List2.ListCount <> 0 Then
Form2.Show
End If
End Function
Function FindFilesAPIcloudq(path As String, SearchStr As String, _
FileCount As Integer, DirCount As Integer) 'if command2 click
Dim FileName As String ' Walking filename variable...
Dim DirName As String ' SubDirectory Name
Dim dirNames() As String ' Buffer for directory name entries
Dim nDir As Integer ' Number of directories in this path
Dim i As Integer ' For-loop counter...
Dim hSearch As Long ' Search Handle
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer
Dim FT As FILETIME
Dim ST As SYSTEMTIME
Dim DateCStr As String, DateMStr As String
If Right(path, 1) <> "\" Then path = path & "\"
' Search for subdirectories.
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(path & "*", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
Do While Cont
DirName = StripNulls(WFD.cFileName)
' Ignore the current and encompassing directories.
If (DirName <> ".") And (DirName <> "..") Then
' Check for directory with bitwise comparison.
If GetFileAttributes(path & DirName) And _
FILE_ATTRIBUTE_DIRECTORY Then
dirNames(nDir) = DirName
DirCount = DirCount + 1
nDir = nDir + 1
ReDim Preserve dirNames(nDir)
' Uncomment the next line to list directories
'List1.AddItem path & FileName
End If
End If
Cont = FindNextFile(hSearch, WFD) ' Get next subdirectory.
Loop
Cont = FindClose(hSearch)
End If
' Walk through this directory and sum file sizes.
hSearch = FindFirstFile(path & SearchStr, WFD)
Cont = True
If hSearch <> INVALID_HANDLE_VALUE Then
While Cont
FileName = StripNulls(WFD.cFileName)
If (FileName <> ".") And (FileName <> "..") And _
((GetFileAttributes(path & FileName) And _
FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
FindFilesAPIcloudq = FindFilesAPIcloudq + (WFD.nFileSizeHigh * _
MAXDWORD) + WFD.nFileSizeLow
FileCount = FileCount + 1
' To list files w/o dates, uncomment the next line
' and remove or Comment the lines down to End If
'List1.AddItem path & FileName
' Include Creation date...
FileTimeToLocalFileTime WFD.ftCreationTime, FT
FileTimeToSystemTime FT, ST
DateCStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
" " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
' and Last Modified Date
FileTimeToLocalFileTime WFD.ftLastWriteTime, FT
FileTimeToSystemTime FT, ST
DateMStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
" " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
List1.AddItem path & FileName & vbTab & _
Format(DateCStr, "mm/dd/yyyy hh:nn:ss") _
& vbTab & Format(DateMStr, "mm/dd/yyyy hh:nn:ss")
MD5File path & FileName
DoEvents
If cloudr(GetMD5Text()) = 3 Then
DeleteFile path & FileName
End If
End If
Cont = FindNextFile(hSearch, WFD) ' Get next file
Wend
Cont = FindClose(hSearch)
End If
' If there are sub-directories...
If nDir > 0 Then
' Recursively walk into them...
For i = 0 To nDir - 1
FindFilesAPIcloudq = FindFilesAPIcloudq + FindFilesAPIcloudq(path & dirNames(i) _
& "\", SearchStr, FileCount, DirCount)
Next i
End If
End Function
Function FindFilesAPIcloud2(path As String, SearchStr As String, _
FileCount As Integer, DirCount As Integer) 'if command2 click
Dialog.Label2.Caption = "����ʹ����ɨ��"
Dim FileName As String ' Walking filename variable...
Dim DirName As String ' SubDirectory Name
Dim dirNames() As String ' Buffer for directory name entries
Dim nDir As Integer ' Number of directories in this path
Dim i As Integer ' For-loop counter...
Dim hSearch As Long ' Search Handle
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer
Dim FT As FILETIME
Dim ST As SYSTEMTIME
Dim DateCStr As String, DateMStr As String
If Right(path, 1) <> "\" Then path = path & "\"
' Search for subdirectories.
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(path & "*", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
Do While Cont
DirName = StripNulls(WFD.cFileName)
' Ignore the current and encompassing directories.
If (DirName <> ".") And (DirName <> "..") Then
' Check for directory with bitwise comparison.
If GetFileAttributes(path & DirName) And _
FILE_ATTRIBUTE_DIRECTORY Then
dirNames(nDir) = DirName
DirCount = DirCount + 1
nDir = nDir + 1
ReDim Preserve dirNames(nDir)
' Uncomment the next line to list directories
'List1.AddItem path & FileName
End If
End If
Cont = FindNextFile(hSearch, WFD) ' Get next subdirectory.
Loop
Cont = FindClose(hSearch)
End If
' Walk through this directory and sum file sizes.
hSearch = FindFirstFile(path & SearchStr, WFD)
Cont = True
If hSearch <> INVALID_HANDLE_VALUE Then
While Cont
FileName = StripNulls(WFD.cFileName)
If (FileName <> ".") And (FileName <> "..") And _
((GetFileAttributes(path & FileName) And _
FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
FindFilesAPIcloud2 = FindFilesAPIcloud2 + (WFD.nFileSizeHigh * _
MAXDWORD) + WFD.nFileSizeLow
FileCount = FileCount + 1
' To list files w/o dates, uncomment the next line
' and remove or Comment the lines down to End If
'List1.AddItem path & FileName
' Include Creation date...
FileTimeToLocalFileTime WFD.ftCreationTime, FT
FileTimeToSystemTime FT, ST
DateCStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
" " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
' and Last Modified Date
FileTimeToLocalFileTime WFD.ftLastWriteTime, FT
FileTimeToSystemTime FT, ST
DateMStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
" " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
List1.AddItem path & FileName & vbTab & _
Format(DateCStr, "mm/dd/yyyy hh:nn:ss") _
& vbTab & Format(DateMStr, "mm/dd/yyyy hh:nn:ss")
Dialog.Label1.Caption = "Scaning ��" & path & FileName
MD5File path & FileName
DoEvents
If cloudr(GetMD5Text()) = 3 Then
Form2.List1.AddItem path & FileName
End If
End If
Cont = FindNextFile(hSearch, WFD) ' Get next file
Wend
Cont = FindClose(hSearch)
End If
' If there are sub-directories...
If nDir > 0 Then
' Recursively walk into them...
For i = 0 To nDir - 1
FindFilesAPIcloud2 = FindFilesAPIcloud2 + FindFilesAPIcloud2(path & dirNames(i) _
& "\", SearchStr, FileCount, DirCount)
Next i
End If
If List2.ListCount <> 0 Then
Form2.Show
End If
End Function

