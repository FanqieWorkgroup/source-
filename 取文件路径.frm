VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form form1 
   BackColor       =   &H8000000C&
   Caption         =   "番茄U盘保护软件"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8265
   LinkTopic       =   "番茄U盘保护软件"
   ScaleHeight     =   6600
   ScaleWidth      =   8265
   StartUpPosition =   3  '窗口缺省
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
      Caption         =   "关于"
      Height          =   2295
      Left            =   1080
      TabIndex        =   3
      Top             =   3840
      Width           =   6735
      Begin VB.CommandButton Command4 
         Caption         =   "联系我们"
         Height          =   975
         Left            =   3720
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         Caption         =   "关于软件"
         Height          =   1095
         Left            =   480
         TabIndex        =   7
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "U盘保护"
      Height          =   3855
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "主动防护设置"
         Height          =   1575
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "优盘再扫描"
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






Private Sub Command1_Click()
frmOptions.Show
End Sub

Private Sub Command2_Click()

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
FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
Screen.MousePointer = vbDefault
Next i
End Sub


 Function USBDISK()
Dim i As Long
    For i = Asc("A") To Asc("Z")
        If GetDriveType(Chr(i) + ":") = 2 Then List3.AddItem Chr(i)
    Next i
End Function



'引用microsoft scprting runtime

'窗体代码

Function FindFilesAPI(path As String, SearchStr As String, _
FileCount As Integer, DirCount As Integer)
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
End Function





Private Sub Command3_Click()
MsgBox "番茄优盘保护软件 V1.0正式版    病毒库版本：db.2018051jf", , "关于"
End Sub

Private Sub Command4_Click()
MsgBox "bug反馈：835078903@qq.com,合作：qq同步,官网：fanqie.gq,更新地址：fanqiesupportpage.gq", , "联系我们"


End Sub

Private Sub Form_Load()
frmOptions.Check1.Value = 1
Register (App.path & "\sqlite3.dll")
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
FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
Screen.MousePointer = vbDefault
Next i
Else
End If
End Sub
