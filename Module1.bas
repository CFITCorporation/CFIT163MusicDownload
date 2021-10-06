Attribute VB_Name = "Module1"
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 = 65001
Function Utf8ToUnicode(ByRef Utf() As Byte) As String
Dim lRet As Long
Dim lLength As Long
Dim lBufferSize As Long
lLength = UBound(Utf) - LBound(Utf) + 1
If lLength <= 0 Then Exit Function
lBufferSize = lLength * 2
Utf8ToUnicode = String$(lBufferSize, Chr(0))
lRet = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf(0)), lLength, StrPtr(Utf8ToUnicode), lBufferSize)
If lRet <> 0 Then
Utf8ToUnicode = Left(Utf8ToUnicode, lRet)
Else
Utf8ToUnicode = ""
End If
End Function
