Attribute VB_Name = "modUtf8"
Option Explicit

' =============================================================================
' modUtf8
' VBAのUnicode文字列と、HTTPで一般的に使われるUTF-8のバイト列を変換します。
'
' VBAのStringは内部でUnicodeとして保持されます。一方、Winsockのsend/recvが
' 扱うのは「文字列」ではなく「バイト列」です。この違いを曖昧にすると、
' 日本語が文字化けしたり、Content-Lengthが一致しなくなったりします。
' =============================================================================

Private Const CP_UTF8 As Long = 65001

Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As LongPtr, _
    ByVal lpUsedDefaultChar As LongPtr) As Long

Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long) As Long

' Unicode文字列をUTF-8バイト列へ変換します。
' 戻り値にはBOM（EF BB BF）を付けません。HTTP本文では通常BOMを不要とするためです。
Public Function Utf8_Encode(ByVal value As String) As Byte()
    Dim result() As Byte
    Dim requiredBytes As Long
    Dim convertedBytes As Long

    If Len(value) = 0 Then
        Utf8_Encode = result
        Exit Function
    End If

    requiredBytes = WideCharToMultiByte( _
        CP_UTF8, 0, StrPtr(value), Len(value), 0, 0, 0, 0)

    If requiredBytes <= 0 Then
        Err.Raise vbObjectError + 2100, "Utf8_Encode", _
            "UTF-8への変換に必要なバイト数を取得できませんでした。"
    End If

    ReDim result(0 To requiredBytes - 1)
    convertedBytes = WideCharToMultiByte( _
        CP_UTF8, 0, StrPtr(value), Len(value), _
        VarPtr(result(0)), requiredBytes, 0, 0)

    If convertedBytes <> requiredBytes Then
        Err.Raise vbObjectError + 2101, "Utf8_Encode", _
            "UTF-8への変換に失敗しました。"
    End If

    Utf8_Encode = result
End Function

' UTF-8バイト列をVBAのUnicode文字列へ変換します。
' byteCountを別引数にしているのは、受信配列の確保サイズではなく、
' 実際にrecvで受信したバイト数だけを変換するためです。
Public Function Utf8_Decode(ByRef bytes() As Byte, ByVal byteCount As Long) As String
    Dim result As String
    Dim requiredCharacters As Long
    Dim convertedCharacters As Long

    If byteCount <= 0 Then
        Utf8_Decode = vbNullString
        Exit Function
    End If

    requiredCharacters = MultiByteToWideChar( _
        CP_UTF8, 0, VarPtr(bytes(LBound(bytes))), byteCount, 0, 0)

    If requiredCharacters <= 0 Then
        Err.Raise vbObjectError + 2102, "Utf8_Decode", _
            "UTF-8の文字数を取得できませんでした。"
    End If

    result = String$(requiredCharacters, vbNullChar)
    convertedCharacters = MultiByteToWideChar( _
        CP_UTF8, 0, VarPtr(bytes(LBound(bytes))), byteCount, _
        StrPtr(result), requiredCharacters)

    If convertedCharacters <> requiredCharacters Then
        Err.Raise vbObjectError + 2103, "Utf8_Decode", _
            "UTF-8からUnicodeへの変換に失敗しました。"
    End If

    Utf8_Decode = result
End Function

' 未初期化の動的配列にUBoundを使うとエラーになるため、
' 安全にバイト数を得る共通関数を用意しています。
Public Function Bytes_Length(ByRef bytes() As Byte) As Long
    On Error GoTo EmptyArray

    Bytes_Length = UBound(bytes) - LBound(bytes) + 1
    Exit Function

EmptyArray:
    Bytes_Length = 0
End Function
