Attribute VB_Name = "ModByteToUTF"
Option Explicit
  
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 = 65001
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
  
Public Function UTF8_Encode(ByVal strUnicode As String) As Byte()
'UTF-8 编码
  
    Dim TLen As Long
    Dim lngBufferSize As Long
    Dim lngResult As Long
    Dim bytUtf8() As Byte
      
    TLen = Len(strUnicode)
    If TLen = 0 Then Exit Function
      
    lngBufferSize = TLen * 3 + 1
    ReDim bytUtf8(lngBufferSize - 1)
      
    lngResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strUnicode), TLen, bytUtf8(0), lngBufferSize, vbNullString, 0)
      
    If lngResult <> 0 Then
        lngResult = lngResult - 1
        ReDim Preserve bytUtf8(lngResult)
    End If
      
    UTF8_Encode = bytUtf8
End Function
Public Function UTF8_Decode(ByRef bUTF8() As Byte) As String
'UTF-8 解码
    Dim lRet As Long
    Dim lLen As Long
    Dim lBufferSize As Long
    Dim sBuffer As String
    Dim bBuffer() As Byte
      
    lLen = UBound(bUTF8) + 1
      
    If lLen = 0 Then Exit Function
      
    lBufferSize = lLen * 2
      
    sBuffer = String$(lBufferSize, Chr(0))
      
    lRet = MultiByteToWideChar(CP_UTF8, 0, VarPtr(bUTF8(0)), lLen, StrPtr(sBuffer), lBufferSize)
      
    If lRet <> 0 Then
        sBuffer = Left(sBuffer, lRet)
    End If
      
    UTF8_Decode = sBuffer
End Function
  
Public Function CreateStringFromByte(ByRef byteArray() As Byte, ByVal ByteLength As Long) As String
'字节数组中的数据连接成字符串
  
    Dim StringData As String
      
    '** 分配字符串空间
    StringData = Space(ByteLength)
    '** 复制字符数组地址内容到字符串地址
    MoveMemory ByVal StringData, ByVal VarPtr(byteArray(0)), ByteLength
      
    '** 返回字符串
    CreateStringFromByte = StringData
End Function
