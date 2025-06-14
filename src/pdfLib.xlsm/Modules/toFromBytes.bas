Attribute VB_Name = "toFromBytes"
' helper functions to convert raw bytes to/from VBA String values
Option Explicit

#Const FASTCOPY_ = True

#If FASTCOPY_ Then
' void RtlMoveMemory(_Out_ void *Destination, _In_ const void *Source, _In_ SIZE_T Length);
'Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)
#End If


' Converts a Byte() to a Unicode (UTF-16 / wide character) string
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
' Converts a Unicode (UTF-16 / wide character) string to a Byte()
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Private Enum CodePages
    CP_UTF8 = 65001
End Enum


' Returns size (count of bytes) of a Byte() array
' Note: returns -1 if unitialized
Function ByteArraySize(bytes() As Byte) As Long
    ' if array has been erased then any access to UBound will raise an Error
    ' but if array is declared but not yet dimension, LBound() is 0 and UBound() is -1
    On Error GoTo errHandler
    If UBound(bytes) >= LBound(bytes) Then ' if not dim'd compares -1 to 0
        ByteArraySize = UBound(bytes) - LBound(bytes)
    End If
    Exit Function
errHandler:
    ByteArraySize = -1
End Function




' returns index into Byte() array that token starts at or -1 if no match
Function FindToken(ByRef buffer() As Byte, ByRef token As String, Optional startAt As Long = -1, Optional searchBackward As Boolean = True) As Long
    On Error GoTo noMatch
    FindToken = -1 ' default to not found
    
    Dim ndx As Long, i As Long
    Dim tokenB() As Byte
    Dim tokenLen As Long
    Dim upperBound As Long: upperBound = ByteArraySize(buffer)
    If upperBound < 0 Then Exit Function ' no data, so token can't be found
    
    ' set startAt index if not explicitly provided
    If (startAt < LBound(buffer)) Or (startAt > UBound(buffer)) Then
        If searchBackward Then
            startAt = UBound(buffer)
        Else
            startAt = LBound(buffer)
        End If
    End If
    
    tokenB = StringToBytes(token)
    tokenLen = UBound(tokenB) - LBound(tokenB) + 1
    If searchBackward Then ' match requires at least tokenLen bytes
        For ndx = startAt - tokenLen To LBound(buffer) Step -1
            For i = 0 To tokenLen - 1
                If buffer(ndx + i) <> tokenB(i) Then GoTo noMatchBackwards
                DoEvents
            Next i
            GoTo FoundMatch
noMatchBackwards:
        Next ndx
    Else
        For ndx = startAt To upperBound - tokenLen
            For i = 0 To tokenLen - 1
                If buffer(ndx + i) <> tokenB(i) Then GoTo noMatchForward
                DoEvents
            Next i
            GoTo FoundMatch
noMatchForward:
        Next ndx
    End If

noMatch:
    ' no match, return -1
    ndx = -1
FoundMatch: ' found a match return ndx it starts at
    FindToken = ndx
End Function


' returns as String all bytes from an offset until vbLF character is found, i.e. returns from offset to end of the line
' offset is updated to index of linefeed character (or last character)
Function GetLine(ByRef bytes() As Byte, ByRef offset As Long) As String
    On Error GoTo errHandler
    Dim i As Long
    For i = offset To UBound(bytes)
        If bytes(i) = 10 Then Exit For
    Next i
    GetLine = BytesToString(bytes, offset, i - offset)
    offset = i
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Function


' returns as String all bytes from an offset until whitespace or non-regular character is found, i.e. returns from offset to end of current token
' offset is updated to index of whitespace character (or last character)
Function GetWord(ByRef bytes() As Byte, ByRef offset As Long) As String
    On Error GoTo errHandler
    Dim i As Long
    For i = offset To UBound(bytes)
        'If IsWhiteSpace(Chr(bytes(i))) Then Exit For
        Select Case bytes(i)
            Case 32, 10, 13, 9, 12 ' " " \n \r \t \f
                Exit For
            Case 47, 40, 41, 91, 93 ' / ( ) [ ]
                If i <> offset Then Exit For   ' if this is first character, then consider it part of same word not starting new word
            Case 60, 62 ' < or <<, > or >>
                ' is this first character and thus part of word or is this ending a word?
                If i = offset Then
                    If bytes(i + 1) = bytes(i) Then ' << or >>
                        i = i + 2
                    Else ' < or >
                        i = i + 1
                    End If
                End If
                ' word ends regardless
                Exit For
        End Select
        DoEvents
    Next i
    GetWord = BytesToString(bytes, offset, i - offset)
    offset = i
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Function


' updates offset to start of next non-whitespace index
' Note: also skips any comments %<blah>EOL
Function SkipWhiteSpace(ByRef bytes() As Byte, ByVal offset As Long, Optional ByVal skipComments As Boolean = True) As Long
    Do While True
        If offset > UBound(bytes) Then Exit Do
        Select Case bytes(offset)
            Case 32, 10, 13, 9, 12
                ' do nothing, keep looping
            Case 37 '%
                If skipComments Then
                    ' loop until we reach a end of line <NL>
                    Do While bytes(offset) <> 10
                        offset = offset + 1
                        If offset > UBound(bytes) Then Exit Do
                        DoEvents
                    Loop
                Else ' not a whitespace so done
                    Exit Do
                End If
            Case Else
                ' not a whitespace character, so done skipping
                Exit Do
        End Select
        offset = offset + 1
        DoEvents
    Loop

    SkipWhiteSpace = offset
End Function


' Helper function to copy range of bytes from one array to another
Sub CopyBytes(ByRef src() As Byte, ByRef dst() As Byte, _
                   Optional ByVal srcOffset As Long = 0, Optional ByVal dstOffset As Long = 0, _
                   Optional ByVal Count As Long = -1 _
                  )
    ' exit early if nothing to copy, avoids out of bounds for FASTCOPY if dstOffset = UBound(dst)+1 and count=0
    If Count = 0 Then Exit Sub
    If srcOffset < LBound(src) Then srcOffset = LBound(src)
    If dstOffset < LBound(dst) Then dstOffset = LBound(dst)
    If (Count < 0) Or ((Count + srcOffset) > (UBound(src) + 1)) Then Count = UBound(src) - srcOffset + 1
    If dstOffset + Count > UBound(dst) + 1 Then ReDim Preserve dst(LBound(dst) To (dstOffset + Count - 1))
#If FASTCOPY_ Then
    CopyMemory VarPtr(dst(dstOffset)), VarPtr(src(srcOffset)), Count
#Else
    Dim i As Long
    For i = 0 To Count - 1
        dst(dstOffset + i) = src(srcOffset + i)
        DoEvents
    Next i
#End If
End Sub


' Helper function to convert byte array to string
Function BytesToString(ByRef bytes() As Byte, Optional offset As Long = 0, Optional Count As Long = -1) As String
    Dim i As Long
    Dim tempStr As String
    If offset < LBound(bytes) Then offset = LBound(bytes)
    If Count < 0 Then Count = UBound(bytes) + 1 - offset
    If (Count + offset) > (UBound(bytes) + 1) Then Count = UBound(bytes) + 1 - offset
    tempStr = Space(Count)
    For i = offset To offset + Count - 1
        Mid(tempStr, i - offset + 1, 1) = Chr(bytes(i))
        DoEvents
    Next i
    BytesToString = tempStr
End Function


''' Return VBA "Unicode" string from byte array encoded in UTF-8
' From: https://www.di-mgt.com.au/howto-convert-vba-unicode-to-utf8.html
Public Function Utf8BytesToString(ByRef abUtf8Array() As Byte, Optional ByVal nBytes As Long = -1) As String
    Dim nChars As Long
    Dim strOut As String
    Utf8BytesToString = ""
    ' Catch uninitialized input array
    If nBytes <= 0 Then nBytes = ByteArraySize(abUtf8Array)
    If nBytes <= 0 Then Exit Function
    ' Get number of characters in output string
    nChars = MultiByteToWideChar(CodePages.CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, 0&, 0&)
    ' Dimension output buffer to receive string
    strOut = String(nChars, 0)
    nChars = MultiByteToWideChar(CodePages.CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, StrPtr(strOut), nChars)
    Utf8BytesToString = Left$(strOut, nChars)
End Function


' Helper function to convert string to byte array (0..N-1)
Function StringToBytes(str As String) As Byte()
    If Len(str) < 1 Then
        StringToBytes = vbNullString
    Else
        StringToBytes = StrConv(str, vbFromUnicode) ' Note only low bytes, mangles any Unicode character > 255, same as Asc() to Byte would
    End If
End Function


' Helper to add a UTF-8 BOM to beginning of a Byte() array
Function AddUtf8BOM(ByRef bytes() As Byte) As Byte()
    Dim ucBytes() As Byte
    ReDim ucBytes(0 To UBound(bytes) + 3) ' +3 for UTF-8 BOM
    ' UTF-8 BOM  239, 187 and 191 in decimal (357, 273, 277 in octal and EF, BB, BF hexadecimal).
    ucBytes(0) = 239: ucBytes(1) = 187: ucBytes(2) = 191
    CopyBytes bytes, ucBytes, 0, 3
    AddUtf8BOM = ucBytes
End Function
    
''' Return byte array with VBA "Unicode" string encoded in UTF-8
' Optionally include UTF-8 Byte Order Mark
' From: https://www.di-mgt.com.au/howto-convert-vba-unicode-to-utf8.html
Public Function StringToUtf8Bytes(ByRef strInput As String, Optional ByVal BOM As Boolean = False) As Byte()
    Dim nBytes As Long
    Dim abBuffer() As Byte
    ' Catch empty or null input string
    StringToUtf8Bytes = vbNullString
    If Len(strInput) < 1 Then Exit Function
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CodePages.CP_UTF8, 0&, ByVal StrPtr(strInput), -1, 0&, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need (0 to (nBytes-1)-1)
    nBytes = WideCharToMultiByte(CodePages.CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    If BOM Then abBuffer = AddUtf8BOM(abBuffer)
    StringToUtf8Bytes = abBuffer
End Function


' left shift bits filling with 0s from the right, note value is capped to 32 bit value
' Warning: this requires LongLong support as written (64bit Excel) to avoid overflow issues
Function LShift(ByVal Value As Long, ByVal shift As Integer) As Long
'    Debug.Print "Value=" & Hex(value) & " << " & shift
    'LShift = CLng((Value * (2^ ^ shift)) And &H7FFFFFFF^)  ' truncate if exceeds max positive Long value
    
    If shift > 31 Then
        LShift = 0
    Else
        Dim fullValue As LongLong
        fullValue = (Value * (2^ ^ shift)) And &HFFFFFFFF^  ' this still sign extends
        'LShift = CLng(fullValue)   ' Will cause Overflow error if > &H7FFFFFFF^
        Dim signBit As Boolean
        signBit = Sgn(fullValue) < 0
        LShift = CLng(fullValue And &H7FFFFFFF^)
        If signBit Then LShift = LShift Or &H80000000
    End If
    
'    Debug.Print Hex(LShift)
    DoEvents
End Function

' right shift bits filling with 0s from the left, result is always 0 or smaller than initial value
' Warning: this requires LongLong support as written (64bit Excel) to avoid overflow issues
Function RShift(ByVal Value As Long, ByVal shift As Integer) As Long
'    Debug.Print "Value=" & Hex(value) & " >> " & shift
    If shift > 31 Then
        RShift = 0
    Else
        RShift = CLng(Int((Value And &HFFFFFFFF^) / (2 ^ shift)) And &H7FFFFFFF^)
    End If
'    Debug.Print Hex(RShift)
    DoEvents
End Function

' Returns the next bitCount bits from the Byte buffer, assuming byteOffset & bitOffset are where to start, updates as bits read
' Note: bitCount should be <= 24  [31 - Max(bitOffset)]
Function nBITS(ByVal bitCount As Integer, ByRef bytes() As Byte, ByRef byteOffset As Long, ByRef bitOffset) As Long
    ' treat 0 bitCount as a NOP
    If bitCount <= 0 Then Exit Function ' nBITS=0

    ' if bitOffset is larger than 1 byte, then adjust byteOffset accordingly so bitOffset < 8
    If bitOffset > 7 Then
        byteOffset = byteOffset + (bitOffset \ 8)
        bitOffset = bitOffset Mod 8
    End If

    ' get next 4 bytes from buffer, but if exceeds bounds use 0 as filler
    Dim b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte
    If UBound(bytes) >= byteOffset Then
        b0 = bytes(byteOffset)
    Else
        b0 = 0
    End If
    If UBound(bytes) >= (byteOffset + 1) Then
        b1 = bytes(byteOffset + 1)
    Else
        b1 = 0
    End If
    If UBound(bytes) >= (byteOffset + 2) Then
        b2 = bytes(byteOffset + 2)
    Else
        b2 = 0
    End If
    If UBound(bytes) >= (byteOffset + 3) Then
        b3 = bytes(byteOffset + 3)
    Else
        b3 = 0
    End If
    
    Dim dw As Long  ' up to 32 bits from bytes buffer
    dw = LShift(b3, 24) Or LShift(b2, 16) Or LShift(b1, 8) Or b0
    
    ' shift to right any bits to ignore, filling left with 0 bits
    dw = RShift(dw, bitOffset)
    
    ' mask of any high order bits still left that we want to ignore as well
    Dim mask As Long
    mask = RShift(-1, 32 - bitCount) ' get value with bitCount 1 bits, -1=1111..1111b >> (32-bitCount) ==> 00..11b
    
    nBITS = dw And mask
    
    bitOffset = bitOffset + bitCount
    ' if bitOffset is larger than 1 byte, then adjust byteOffset accordingly so bitOffset < 8
    If bitOffset > 7 Then
        byteOffset = byteOffset + (bitOffset \ 8)
        bitOffset = bitOffset Mod 8
    End If
End Function



' pads a 32 bit hex number with leading 0s
Function HexStr(ByVal v As LongLong) As String
    Const pad As String = "00000000"
    Dim s As String: s = CStr(Hex(v))
    Dim l As Long: l = Len(s)
    If l >= 8 Then
        HexStr = s
    Else
        HexStr = Left(pad, 8 - Len(s)) & s
    End If
End Function



' reads in contents of filename returning as a Byte() array
' fileLen is set to size in bytes of the file
Function readFile(ByVal filename As String, ByRef fileLen As Long) As Byte()
    On Error GoTo errHandler
    Dim fileNum As Integer
    Dim content() As Byte
    
    ' Open file and determine size
    fileNum = FreeFile
    Open filename For Binary Access Read Shared As #fileNum
    fileLen = LOF(fileNum)
    ' arbitrarily limit files to 512MB
    If fileLen > 512& * 1024& * 1024& Then
        'MsgBox "Warning exceeds supported file size!", vbOKOnly Or vbCritical, filename
        Debug.Print "Warning exceeds supported file size! " & filename & " is " & (fileLen + 524288) \ (1024& * 1024&) & "MB"
        fileLen = -1 * fileLen
        Exit Function
    End If
    
    ' allocate memory and read in data
    If fileLen > 0 Then ReDim content(fileLen - 1)
    Get #fileNum, , content
    
cleanup:
    On Error Resume Next
    Close #fileNum
    readFile = content
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    fileLen = 0
    ReDim content(0)
    Stop
    Resume cleanup
End Function


' writes string to file as bytes and returns count of bytes written
Function PutString(ByVal fileNum As Integer, ByRef str As String) As Long
    Dim data() As Byte: data = StringToBytes(str)
    PutString = PutBytes(fileNum, data)
End Function


' writes bytes to file and returns count of bytes written
Function PutBytes(ByVal fileNum As Integer, ByRef data() As Byte) As Long
    On Error GoTo errHandler
    Dim byteCount As Long
    byteCount = UBound(data) - LBound(data) + 1
    Put #fileNum, , data
    PutBytes = byteCount
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function
