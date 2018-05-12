Attribute VB_Name = "Module2"
Option Explicit
Type MD5_CTX
      dwNUMa      As Long
      dwNUMb      As Long
      Buffer(15)  As Byte
      cIN(63)     As Byte
      cDig(15)    As Byte
End Type
 
Private Declare Sub MD5Init Lib "advapi32" (lpContext As MD5_CTX)
Private Declare Sub MD5Final Lib "advapi32" (lpContext As MD5_CTX)
Private Declare Sub MD5Update Lib "advapi32" (lpContext As MD5_CTX, _
                           ByRef lpBuffer As Any, ByVal BufSize As Long)
 
Private stcContext   As MD5_CTX
Public Function MD5File(ByVal FileName As String) As Byte()
   Const BUFFERSIZE  As Long = 1024& * 512      ' 缓冲区 512KB
      Dim DataBuff() As Byte
      Dim lFileSize  As Long
      Dim iFn        As Long
 
   On Error GoTo E_Handle_MD5
   If (Len(Dir$(FileName)) = 0) Then Err.Raise 5      '文件不存在
   ReDim DataBuff(BUFFERSIZE - 1)
   iFn = FreeFile()
   Open FileName For Binary As #iFn
   lFileSize = LOF(iFn)
   Call MD5Init(stcContext)
   If (lFileSize = 0) Then
      Call MD5Update(stcContext, 0, 0)
   Else
      Do While (lFileSize > 0)
         Get iFn, , DataBuff
         If (lFileSize > BUFFERSIZE) Then
            Call MD5Update(stcContext, DataBuff(0), BUFFERSIZE)
         Else
            Call MD5Update(stcContext, DataBuff(0), lFileSize)
         End If
         lFileSize = lFileSize - BUFFERSIZE
      Loop
   End If
   Close iFn
   Call MD5Final(stcContext)
E_Handle_MD5:
   MD5File = stcContext.cDig
End Function
Public Function GetMD5Text() As String
      Dim sResult As String, i&
   If (stcContext.dwNUMa = 0) Then
      sResult = vbNullString
   Else
      sResult = Space$(32)
      For i = 0 To 15
         Mid$(sResult, i + i + 1) = Right$("0" & Hex$(stcContext.cDig(i)), 2)
      Next
   End If
   GetMD5Text = LCase$(sResult)       ' LCase$(sResult) '字母小写
End Function
Public Function MD5Bytes(Buffer() As Byte, _
                        Optional ByVal size As Long = -1) As Byte()
      Dim U As Long, pBase   As Long
 
   pBase = LBound(Buffer)
   U = UBound(Buffer) - pBase
   If (-1 = size) Then size = U + 1
   Call MD5Init(stcContext)
   If (-1 = U) Then
      Call MD5Update(stcContext, 0, 0)
   Else
      Call MD5Update(stcContext, Buffer(pBase), size)
   End If
   Call MD5Final(stcContext)
   MD5Bytes = stcContext.cDig
End Function
Public Function MD5String(strText As String) As Byte()
      Dim aBuffer()     As Byte
 
   Call MD5Init(stcContext)
   If (Len(strText) > 0) Then
      aBuffer = StrConv(strText, vbFromUnicode)
      Call MD5Update(stcContext, aBuffer(0), UBound(aBuffer) + 1)
   Else
      Call MD5Update(stcContext, 0, 0)
   End If
   Call MD5Final(stcContext)
   MD5String = stcContext.cDig
End Function
 
 






