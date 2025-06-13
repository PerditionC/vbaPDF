Attribute VB_Name = "pdfParseAndGetValues"
' parses and laods PDF values
Option Explicit

Enum PDF_ValueType
    PDF_Null = 0
    PDF_Name
    PDF_Boolean
    PDF_Integer
    PDF_Real
    PDF_String
    PDF_Array
    PDF_Dictionary
    PDF_Stream      ' actual stream object with dictionary and data
    PDF_StreamData  ' represents only stream ... endstream portion
    
    ' to simplify processing, not one of the 9 basic types either
    PDF_Object      ' id generation obj << dictionary >> endobj
    PDF_Reference   ' indirect object
    PDF_Comment     ' not used as comments skipped along with whitespace
    PDF_Trailer
    
    ' markers, no actual value returned
    PDF_EndOfArray
    PDF_EndOfDictionary
    PDF_EndOfStream
    PDF_EndOfObject
End Enum


' given raw PDF file contents as Byte array and offset in the array to peak, returns value type at offset
Function GetValueType(ByRef bytes() As Byte, ByVal offset As Long) As PDF_ValueType
    On Error GoTo errHandler
    GetValueType = PDF_ValueType.PDF_Null
    
    offset = SkipWhiteSpace(bytes, offset)
    If offset > UBound(bytes) Then Exit Function    ' return null type if end of data
    Dim token As String: token = Chr(bytes(offset))
    
    Dim tmpStr As String
    Select Case LCase(token)
        Case "n"
            If Not IsMatch(GetWord(bytes, offset), "null") Then Stop ' error! expecting null or value
            GetValueType = PDF_ValueType.PDF_Null
        Case "/"
            GetValueType = PDF_ValueType.PDF_Name
        Case "t", "f"
            tmpStr = GetWord(bytes, offset)
            Select Case LCase(tmpStr)
                Case "trailer"
                    GetValueType = PDF_ValueType.PDF_Trailer
                Case "true", "false"
                    GetValueType = PDF_ValueType.PDF_Boolean
                Case Else
                    Stop ' error! unexpected value
            End Select
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "+", "-", "."
            'could also be indirect reference or direct reference (obj)
            'Note: its ok here that offset is wrong if not obj or R as local offset value is discarded and not updated when getting actual value
            Dim words(0 To 2) As String
            words(0) = GetWord(bytes, offset)
            offset = SkipWhiteSpace(bytes, offset)
            words(1) = GetWord(bytes, offset)
            offset = SkipWhiteSpace(bytes, offset)
            words(2) = GetWord(bytes, offset)
            'offset = SkipWhiteSpace(bytes, offset)
            If IsNumeric(words(0)) And IsNumeric(words(1)) And IsMatch(words(2), "R") Then
                GetValueType = PDF_ValueType.PDF_Reference
            ElseIf IsNumeric(words(0)) And IsNumeric(words(1)) And IsMatch(words(2), "obj") Then
                GetValueType = PDF_ValueType.PDF_Object
            Else
                If InStr(1, words(0), ".", vbBinaryCompare) > 0 Then
                    GetValueType = PDF_ValueType.PDF_Real
                Else
                    GetValueType = PDF_ValueType.PDF_Integer
                End If
            End If
        Case "s"
            If Not IsMatch(GetWord(bytes, offset), "stream") Then Stop ' error! unexpected value
            GetValueType = PDF_ValueType.PDF_StreamData
        Case "("
            GetValueType = PDF_ValueType.PDF_String
        Case "["
            GetValueType = PDF_ValueType.PDF_Array
        Case "]"
            GetValueType = PDF_ValueType.PDF_EndOfArray
        Case "<"
            If IsMatch("<", Chr(bytes(offset + 1))) Then
                GetValueType = PDF_ValueType.PDF_Dictionary ' <<...>>
            Else
                GetValueType = PDF_ValueType.PDF_String ' hex string <####>
            End If
        Case "%"
            GetValueType = PDF_ValueType.PDF_Comment
            
        Case "e", ">" '
            tmpStr = GetWord(bytes, offset)
            Select Case LCase(tmpStr)
                Case ">>"
                    GetValueType = PDF_ValueType.PDF_EndOfDictionary
                Case "endstream"
                    GetValueType = PDF_ValueType.PDF_EndOfStream
                Case "endobj"
                    GetValueType = PDF_ValueType.PDF_EndOfObject
                Case Else
                    Stop ' error! expecting null or value
                    GetValueType = PDF_ValueType.PDF_Null
            End Select
            
        Case Else
            GetValueType = PDF_ValueType.PDF_Null
            Stop ' error! unexpected value
    End Select
    
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    'Resume
End Function


' a Name may have value encoded in #00 form, ie # followed by 2 hex digits
Function ProcessName(ByRef name As String) As String
    Dim ndx As Long
    ndx = InStr(1, name, "#", vbBinaryCompare)
    If ndx > 0 Then
        Dim s As String, N As Long
        ProcessName = Left(name, ndx - 1)
        s = Mid(name, ndx + 1, 2) ' get hex digits
        N = CLng("&H" & s) ' convert from hex to Long
        s = ProcessName & Chr(N) & Mid(name, ndx + 3) ' combine with hex digit as character
        ProcessName = ProcessName(s) ' for now just recursively call to handle multiple encoded values, really should be a loop!
    Else
        ProcessName = name
    End If
End Function


' process a byte of string data, handles unicode or pdfDocEncoding
Private Sub UpdateString(ByRef offset As Long, ByVal chrValue As Integer, ByRef tmpStr As String, ByRef strBuffer() As Byte, ByRef strLen As Long)
    Select Case strLen
        Case Is < 0 ' no BOM, just store character asis
            tmpStr = tmpStr & Chr(chrValue)
        Case 1 ' check next character continues BOM
            If (chrValue = &HBB) Or (chrValue = &HFE) Then
                strLen = 2
                strBuffer(1) = chrValue
            Else ' woops no BOM
                tmpStr = Chr(strBuffer(0)) & Chr(chrValue)
                strLen = -1
            End If
        Case 2 ' and is full BOM found
            ' either begins UTF-16 data or 3rd byte of UTF-8 BOM
            If chrValue = &HBF Then
                strLen = 3
                strBuffer(2) = chrValue
            ElseIf strBuffer(0) = &HFF Then
                Stop ' sorry UTF-16 not yest supported :-(
            Else ' woops no BOM
                tmpStr = Chr(strBuffer(0)) & Chr(strBuffer(1)) & Chr(chrValue)
                strLen = -1
            End If
        Case Is >= 3
            ' unicode, just store in our Byte() buffer (note: we overwrite BOM)
            strBuffer(strLen - 3) = chrValue
            strLen = strLen + 1
            ' if UTF-16 then use ChrW (chrValue)
        Case Else ' need to determine if UTF-8 or UTF-16BE BOM or not
            If (chrValue = &HEF) Or (chrValue = &HFE) Then
                strLen = 1
                strBuffer(0) = chrValue
            Else ' no BOM
                strLen = -1 ' flag no BOM
                tmpStr = Chr(chrValue)
            End If
    End Select
    DoEvents
End Sub


' returns a value loaded for a PDF
' updates offset to next non-whitespace byte after this value is loaded
' Note: meta is only used for stream object
Function GetValue(ByRef bytes() As Byte, ByRef offset As Long, Optional Meta As Dictionary = Nothing) As pdfValue
    On Error GoTo errHandler
    DoEvents
    Dim result As pdfValue: Set result = New pdfValue
    Set GetValue = result
    result.id = 0
    result.generation = 0
    result.Value = Empty
    result.valueType = GetValueType(bytes, offset)
    If offset > UBound(bytes) Then Exit Function        ' return null type if end of daa
    
    offset = SkipWhiteSpace(bytes, offset)
    If offset > UBound(bytes) Then Exit Function    ' return null type if end of data
    Dim token As String: token = Chr(bytes(offset))
    
    Dim tmpStr As String
    Dim name As pdfValue, Value As pdfValue
    Select Case result.valueType
        Case PDF_ValueType.PDF_Null
            ' result.value = Empty
            If Not IsMatch(GetWord(bytes, offset), "null") Then
                Stop ' error, expecting "null"
            End If
        Case PDF_ValueType.PDF_Name
            result.Value = ProcessName(GetWord(bytes, offset))
        Case PDF_ValueType.PDF_Boolean
            tmpStr = GetWord(bytes, offset)
            result.Value = CBool(tmpStr)
        Case PDF_ValueType.PDF_Integer
            tmpStr = GetWord(bytes, offset)
            result.Value = CLng(tmpStr)
        Case PDF_ValueType.PDF_Real
            tmpStr = GetWord(bytes, offset)
            result.Value = CDbl(tmpStr)
        Case PDF_ValueType.PDF_String
            ' (...) or <hex digits>
            ' may have a UTF-8 BOM bytes 239, 187 and 191
            ' or Unicode UTF-16 BOM bytes 255, 254
            ' Note: we do not process the language[country] escaped information
            ' if provided, the string until end or next 1B escape code will be in
            ' indicated language (default Latin1 or Document/Page level Lang used otherwise
            ' Format: ESC LANG [COUNTRY] ESC e.g. \033enUS\033USING language escape
            ' and \033en\033USING language escape both indicate the string USING language escape
            ' is en (English) with the first en-US and the later country unspecified.
            ' Per the spec, XX should be used for unknown country or preferable not included.
            ' Language is always 2 characters and if provided country is 2 characters
            ' so either \033XX\033 or \033XXYY\033 where XX=language and YY=country
            '    Case 27    ' Esc 1B, flags encoding language and optionally country codes follow
            '        langCode = tmpStr(offset) & tmpStr(offset + 1))
            '        offset = offset + 2
            '        If Asc(tmpStr(offset)) <> 27 Then
            '            cntryCode = tmpStr(offset) & tmpStr(offset + 1)
            '            offset = offset + 2
            '        End If
            '        If Asc(tmpStr(offset)) = 27 Then
            '            Debug.Print "Text found using " & langCode & cntryCode
            '        Else
            '            Debug.Print "Error! out of sync in text string, expecting Esc to mark end of lang-country specifier!"
            '        End If
            tmpStr = vbNullString
            Dim strBuffer(0 To 65535) As Byte ' max size pdf string value
            Dim strLen As Long: strLen = 0
            If bytes(offset) = 40 Then ' Asc("(")
                offset = offset + 1
                ' TODO only need to escape unbalanced ), so need to keep track of balanced ()
                Do While bytes(offset) <> 41 ' Asc(")")
                    If bytes(offset) = 92 Then ' Asc("\") then escaped value
                        offset = offset + 1
                        Select Case bytes(offset)
                            Case 110 ' Asc("n")
                                strBuffer(strLen) = 10 ' Asc(vbLf)
                                strLen = strLen + 1
                            ' \ immediately followed by newline treated as line continuation & ignored
                            ' we treat \<CR><NL>, \<CR> and \<NL> the same (note \<NL><CR> treated as <CR> )
                            Case 13 ' Asc(vbCr)
                                ' line continuation, ignore line break, <CR> or <CR><NL>
                                If bytes(offset + 1) = 10 Then offset = offset + 1
                            Case 10 ' Asc(vbLf)
                                ' line continuation, ignore line break
                            Case 114 ' Asc("r")
                                strBuffer(strLen) = vbCr
                                strLen = strLen + 1
                            Case 116 ' Asc("t")
                                strBuffer(strLen) = vbTab
                                strLen = strLen + 1
                            Case 102 ' Asc("f")
                                strBuffer(strLen) = &HC  ' formfeed
                                strLen = strLen + 1
                            Case 98  ' Asc("b")
                                strBuffer(strLen) = &H8  ' backspace
                                strLen = strLen + 1
                            Case 92  ' Asc("\")
                                strBuffer(strLen) = 92   ' "\"
                                strLen = strLen + 1
                            Case 41  ' Asc(")")
                                strBuffer(strLen) = 41   ' ")"
                                strLen = strLen + 1
                            Case 40  ' Asc("(")
                                strBuffer(strLen) = 40   ' "("
                                strLen = strLen + 1
                            Case 48 To 57 ' Asc("0") To Asc("9")    ' octal
                                Dim octStr As String: octStr = Chr(bytes(offset))
                                Dim octVal As Long: octVal = 0
                                Dim maxOctStrLen As Integer: maxOctStrLen = 3
                                Do While (octStr >= "0") And (octStr <= "9") And (maxOctStrLen > 0)
                                    maxOctStrLen = maxOctStrLen - 1 ' only 1 to 3 octal digits
                                    octVal = (octVal * 8) + CLng(octStr)
                                    offset = offset + 1
                                    octStr = Chr(bytes(offset))
                                Loop
                                offset = offset - 1 ' so we don't skip a character
                                strBuffer(strLen) = octVal
                                strLen = strLen + 1
                            Case Else
                                ' unknown/unexpected escape value
                                Stop
                        End Select
                    Else
                        strBuffer(strLen) = bytes(offset)
                        strLen = strLen + 1
                    End If
                    DoEvents
                    offset = offset + 1
                Loop
                offset = offset + 1 ' skip past ending ")"
            Else ' < hex encoded string >
                offset = offset + 1
                Do While bytes(offset) <> 62 'Asc(">")
                    ' get 2 hex digits, ignoring whitespace, may end with odd # of hex digits
                    offset = SkipWhiteSpace(bytes, offset)
                    Dim HexStr As String
                    HexStr = Chr(bytes(offset))
                    offset = offset + 1
                    offset = SkipWhiteSpace(bytes, offset)
                    If bytes(offset) <> 62 Then ' Asc(">")
                        HexStr = HexStr & Chr(bytes(offset))
                        offset = offset + 1
                    Else
                        HexStr = HexStr & "0"
                    End If
                    
                    Dim hexValue As Integer: hexValue = CLng("&H" & HexStr)
                    strBuffer(strLen) = hexValue
                    strLen = strLen + 1
                    
                    DoEvents
                Loop
                offset = offset + 1 ' skip past ending ">"
            End If
            ' we now have a byte buffer of values 0-255, we need to determine if
            ' encoding is pdfDocEncoding, UTF-8, or UTF-16 (should be BE but apparently LE also used)
            ' UTF-8 has BOM bytes 239, 187 and 191
            ' Unicode UTF-16 has BOM bytes 255, 254 for BigEndian and 254, 255 for LittleEndian
            ' Note: to simplify and not access bytes outside range we get 1st 3 bytes or leave as 0
            Dim b1 As Byte, b2 As Byte, b3 As Byte
            If strLen >= 1 Then b1 = strBuffer(0)
            If strLen >= 2 Then b2 = strBuffer(1)
            If strLen >= 3 Then b3 = strBuffer(2)
            If (b1 = 239) And (b2 = 187) And (b3 = 191) Then ' UTF-8
                tmpStr = Utf8BytesToString(strBuffer, strLen)
                tmpStr = Mid(tmpStr, 2) ' strip BOM
            ElseIf (b1 = 254) And (b2 = 255) Then ' UTF-16BE
                Dim ndx As Long
                For ndx = 2 To strLen - 1
                    Dim chrVal As Integer
                    chrVal = strBuffer(ndx)
                    If (ndx + 1) <= UBound(strBuffer) Then chrVal = chrVal + (strBuffer(ndx + 1) * 256)
                    tmpStr = tmpStr & ChrW(chrVal)
                Next ndx
            ElseIf (b1 = 255) And (b2 = 254) Then ' UTF-16LE
                tmpStr = strBuffer  ' VBA supports converting unicode byte array to string *** TODO verify this if > ASCII
                tmpStr = Mid(tmpStr, 2) ' strip BOM
            Else ' pdfDocEncoding
                'tmpStr = StrConv(strBuffer, vbUnicode)
                tmpStr = vbNullString
                Dim ndx2 As Long
                For ndx2 = 0 To strLen - 1
                    tmpStr = tmpStr & Chr(strBuffer(ndx2))
                Next
            End If
            result.Value = tmpStr
        Case PDF_ValueType.PDF_Array
            offset = offset + 1 ' skip the opening [
            Dim col As Collection
            Set col = New Collection
            Set result.Value = col
            Set Value = GetValue(bytes, offset)
            Do While Value.valueType <> PDF_ValueType.PDF_EndOfArray
                col.Add Value
                Set Value = GetValue(bytes, offset)
            Loop
            Set col = Nothing
        Case PDF_ValueType.PDF_Dictionary
            offset = offset + 2 ' skip past <<
            Dim dict As Dictionary
            Set dict = New Dictionary
            Set result.Value = dict
            Set name = GetValue(bytes, offset)
            Do While name.valueType = PDF_ValueType.PDF_Name
                Set Value = GetValue(bytes, offset)
                dict.Add CStr(name.Value), Value
                ' get name of next element of dictionary (or end of dictionary marker >> )
                Set name = GetValue(bytes, offset)
            Loop
            Set dict = Nothing
        Case PDF_ValueType.PDF_StreamData
            ' Note: PDF_ValueType.PDF_Stream is handled in PDF_Object
            ' this is only the stream ... endstream portion of a PDF stream
            tmpStr = GetLine(bytes, offset) ' skip stream, note dictionary previous read in should have a /Length value (passed in as meta)
            ' skip new line, data starts immediately after and may containe additional whitespace
            If Chr(bytes(offset)) = vbCr Then offset = offset + 1
            If Chr(bytes(offset)) = vbLf Then offset = offset + 1
            result.valueType = PDF_ValueType.PDF_StreamData
            Dim dataBytes() As Byte
            If Meta Is Nothing Then Set Meta = New Dictionary
            If Meta.Exists("/Length") Then
                Dim byteLen As Long
                If IsObject(Meta.Item("/Length").Value) Then
                    ' Reference to actual /Length value, allowed but not common
                    Stop
                    ' TODO pass we need Dictionary of loaded objects or xrefTable to load value
                Else
                    ' /Length directly in file as number, usual
                    byteLen = CLng(Meta.Item("/Length").Value)
                End If
                ReDim dataBytes(0 To byteLen - 1)
                CopyBytes bytes, dataBytes, offset, 0, byteLen
                offset = offset + byteLen
                offset = SkipWhiteSpace(bytes, offset)
                tmpStr = GetWord(bytes, offset) ' endstream
            Else
                ' missing required key, we fake it and just read until we find an endstream line
                Debug.Print "Warning: PDF noncompliant, steram missing required /Length key"
                Dim dataStr As String
                Do While Not IsMatch(Left(tmpStr, 9), "endstream")
                    dataStr = dataStr + tmpStr
                    tmpStr = GetLine(bytes, offset)
                    ' skip past just NewLine
                    If Chr(bytes(offset)) = vbCr Then
                        tmpStr = tmpStr & vbCr
                        offset = offset + 1
                    End If
                    If Chr(bytes(offset)) = vbLf Then
                        tmpStr = tmpStr & vbLf
                        offset = offset + 1
                    End If
                    DoEvents
                Loop
                ' strip final NewLine before "endstream" if added
                If Right(dataStr, 1) = vbLf Then dataStr = Left(dataStr, Len(dataStr) - 1)
                If Right(dataStr, 1) = vbCr Then dataStr = Left(dataStr, Len(dataStr) - 1)
                dataBytes = StringToBytes(dataStr)
            End If
            result.Value = dataBytes
            If Not IsMatch(Left(tmpStr, 9), "endstream") Then Stop ' error unexpected token found!
            
        Case PDF_ValueType.PDF_Reference, PDF_ValueType.PDF_Object
            Dim words(0 To 2) As String
            words(0) = GetWord(bytes, offset)
            offset = SkipWhiteSpace(bytes, offset)
            words(1) = GetWord(bytes, offset)
            offset = SkipWhiteSpace(bytes, offset)
            words(2) = GetWord(bytes, offset)
            offset = SkipWhiteSpace(bytes, offset)
            If IsMatch(words(2), "obj") Then
                result.id = CLng(words(0))
                Set result.Value = GetValue(bytes, offset)
                Dim endObjOrStream As pdfValue
                If result.Value.valueType = PDF_ValueType.PDF_Dictionary Then
                    Set Meta = result.Value.Value
                Else
                    Set Meta = Nothing
                End If
                Set endObjOrStream = GetValue(bytes, offset, Meta)
                ' if actually a stream obj, a dictionary <<>>stream endstream then load stream data
                If endObjOrStream.valueType = PDF_ValueType.PDF_StreamData Then
                    Dim stream As pdfValue
                    Set stream = New pdfValue
                    stream.valueType = PDF_ValueType.PDF_Stream
                    Dim streamObj As pdfStream
                    Set streamObj = New pdfStream
                    streamObj.Init result.Value, endObjOrStream
                    Set stream.Value = streamObj
                    Set streamObj = Nothing
                    Set result.Value = stream
                    Set stream = Nothing
                    Set endObjOrStream = GetValue(bytes, offset)
                End If
                If endObjOrStream.valueType <> PDF_ValueType.PDF_EndOfObject Then
                    Stop ' error, expected "endobj"
                End If
            Else ' ISMatch(word(2), "R") Then
                result.Value = CLng(words(0))
            End If
            result.generation = CLng(words(1))
        Case PDF_ValueType.PDF_Comment
            result.Value = GetLine(bytes, offset)
        Case PDF_ValueType.PDF_Trailer
            If Not IsMatch(GetWord(bytes, offset), "trailer") Then
                Stop ' error, expected "trailer"
            End If
            Set result.Value = GetValue(bytes, offset)
        Case PDF_ValueType.PDF_EndOfArray, PDF_ValueType.PDF_EndOfDictionary, PDF_ValueType.PDF_EndOfObject, PDF_ValueType.PDF_EndOfStream
            tmpStr = GetWord(bytes, offset) ' skip past end marker
        Case Else
            Stop ' error, unexpected type!
    End Select
    
    ' skip past any trailing whitespace
    offset = SkipWhiteSpace(bytes, offset)
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function


' given raw contents of a PDF file as a Byte array, returns offset in array that cross reference table begins
' Note: we start at end of contents, will work with Linearized PDF's if has compatible trailer pointing to xref stream obj per specification.
' Possible future enhancement: optionally look at beginning of content (1st 1024 bytes) for /Linearized declaration
Function GetXrefOffset(ByRef content() As Byte) As Long
    On Error GoTo errHandler
    Dim offset As Long
    offset = FindToken(content, "startxref", searchBackward:=True)
    offset = SkipWhiteSpace(content, offset + Len("startxref"))
    GetXrefOffset = CLng(GetWord(content, offset))
    ' should immediately be followed by end of file marker, EOF marker is a comment, so we can't GetValue as it will be skipped
    offset = SkipWhiteSpace(content, offset, skipComments:=False)
    If offset > UBound(content) Then Exit Function ' we allow invalid PDF documents that miss %%EOF but actually end after cross reference table
    If Not IsMatch(BytesToString(content, offset, 5), "%%EOF") Then
        MsgBox "PDF document missing %%EOF end of file marker, invalid PDF!", vbCritical Or vbOKOnly, "Warning - invalid document"
    End If
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Function


' trailer, see trailer.value.value.item("/Root") and "/Size"
Function GetTrailer(ByRef content() As Byte) As pdfValue
    On Error GoTo errHandler
    Dim offset As Long: offset = FindToken(content, "trailer", searchBackward:=True)
    If offset < 0 Then
        Set GetTrailer = New pdfValue
        GetTrailer.valueType = PDF_ValueType.PDF_Null
    Else
        Set GetTrailer = GetValue(content, offset)
    End If
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Function


' validates content() probably is a PDF document and returns version from PDF declaration
' returns True if found a valid PDF version, False on any error or invalid version declaration
Function GetPdfHeader(ByRef content() As Byte, ByRef headerVersion As String) As Boolean
    On Error GoTo errHandler
    ' PDF documents should begin with something like %PDF-1.7<whitespace>
    Dim pdfHeaderVersion As String
    Dim unusedOffset As Long
    ' extract initial string in file, reusing GetWord but it should work fine for this purpose
    pdfHeaderVersion = TrimWS(GetWord(content, unusedOffset))
    
    ' validate its in epected format
    If Not IsMatch("%PDF-", Left(pdfHeaderVersion, 5)) Then
        Debug.Print "Likely invalid PDF, file begins with [" & pdfHeaderVersion & "] expecting PDF-#.#"
        Exit Function
    End If
    
    ' return header version
    headerVersion = pdfHeaderVersion
    
    GetPdfHeader = True ' success
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Function


' given a pdfValue dictionary object, returns the values associated with a given name
Function GetDictionaryValue(ByRef Value As pdfValue, ByVal name As String) As pdfValue
    On Error GoTo errHandler
    If Value.Value.Exists(name) Then
        Set GetDictionaryValue = Value.Value.Item(name)
    Else
        Set GetDictionaryValue = New pdfValue   ' defaults to PDF_ValueType.PDF_Null
    End If
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Function


' returns /Root obj in PDF
' expects to be given the pdfValue returned from GetTrailer() call
Function GetRoot(ByRef trailer As pdfValue) As pdfValue
    Set GetRoot = GetDictionaryValue(trailer.Value, "/Root")
End Function


' returns /Info obj in PDF if exists, otherwise return PDF_Null
' expects to be given the pdfValue returned from GetTrailer() call
Function GetInfo(ByRef trailer As pdfValue) As pdfValue
    Set GetInfo = GetDictionaryValue(trailer.Value, "/Info")
End Function


' returns how many xref entries in xref table (/Size value)
' expects to be given the pdfValue returned from GetTrailer() call
Function GetXrefSize(ByRef trailer As pdfValue) As Long
    On Error GoTo errHandler
    Dim Size As pdfValue
    Set Size = GetDictionaryValue(trailer.Value, "/Size")
    If Size.valueType = PDF_ValueType.PDF_Integer Or Size.valueType = PDF_ValueType.PDF_Real Then
        GetXrefSize = CLng(Size.Value)
    End If
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Function


' reads in byteCount bytes and returns value as Long, offset is updated
Private Function getInt(ByRef content() As Byte, ByRef offset As Long, ByVal byteCount As Long, ByVal defaultValue As Long) As Long
    If byteCount < 1 Then
        getInt = defaultValue
    Else
        Dim i As Long, Value As Long
        Value = 0
        For i = 0 To byteCount - 1
            Value = (Value * 256) + content(offset)
            offset = offset + 1
        Next i
        getInt = Value
    End If
End Function


' extract xref table (size # of xref entries)
' Note: trailer may be PDF_Null if cross reference stream, if so updated to xref stream dictionary << >>
' catalog should not be provided (or pass Nothing) to parse primary cross reference table
' and should be a valid Dictionary object for all Prev'ious catalogs parsed
Function ParseXrefTable(ByRef content() As Byte, ByRef offset As Long, ByRef trailer As pdfValue, Optional ByRef xrefTable As Dictionary = Nothing) As Dictionary ' of xrefEntry
    On Error GoTo errHandler
    Dim primaryXref As Boolean ' = False
    If xrefTable Is Nothing Then
        Set xrefTable = New Dictionary
        primaryXref = True
    End If
    Set ParseXrefTable = xrefTable

    Dim Size As Long    ' how many total entries
    Dim Count As Long   ' how many entries in current xref chunk
    Dim id As Long      ' current catalog entry id
    Dim str As String   ' tmp string
    Dim entry As xrefEntry

    ' determine if it's standard cross reference table (starts with xref) or cross reference stream object
    Dim peakOffset As Long: peakOffset = offset  ' temp variable so we don't actually change our offset
    Dim xrefType As String
    xrefType = GetWord(content, peakOffset)
    If IsMatch(xrefType, "xref") Then
        ' skip past "xref"
        GetWord content, offset
    
        offset = SkipWhiteSpace(content, offset)
    
        ' get Size value from trailer, # of expected total entries
        Size = GetXrefSize(trailer)
        
        ' set of xrefs with in form
        ' id count
        ' id/offset generation f/n
        Do While Size > 0
            str = GetWord(content, offset)
            offset = SkipWhiteSpace(content, offset)
            If Not IsNumeric(str) Then Exit Do ' found end of xref (trailer?) but not all xref entries
            id = CLng(str) ' staring id#
            str = GetWord(content, offset)
            offset = SkipWhiteSpace(content, offset)
            Count = CLng(str)
            Dim entryCount As Long: entryCount = Count ' # of entries in this section
            Do While entryCount > 0
                Set entry = New xrefEntry
                entry.id = id
                entry.generation = 0
                entry.isFree = False
                entry.nextFreeId = 0
                entry.offset = 0
            
                ' get id/offset value
                str = GetWord(content, offset)
                offset = SkipWhiteSpace(content, offset)
                entry.offset = CLng(str)
                ' get generation
                str = GetWord(content, offset)
                offset = SkipWhiteSpace(content, offset)
                entry.generation = CLng(str)
                ' get flag if f=free or n=action obj
                str = GetWord(content, offset)
                offset = SkipWhiteSpace(content, offset)
                If IsMatch(str, "f") Then
                    entry.isFree = True
                    entry.nextFreeId = entry.offset
                    entry.offset = 0
                End If
                
                ' add/replace to our catalog
                If xrefTable.Exists(entry.id) Then
                    If (entry.id <> 0) And primaryXref Then
                        Debug.Print "Warning: duplicate obj " & entry.id & " found!"
                        Stop
                    End If
                Else
                    xrefTable.Add entry.id, entry
                    Set entry = Nothing
                End If
            
                id = id + 1
                entryCount = entryCount - 1
                DoEvents
            Loop
            Size = Size - Count
            DoEvents
        Loop
    
        If Size > 0 Then
            Debug.Print "Warning: PDF did not find /Size cross reference entries - missing " & Size & " entries."
        End If
    Else    ' either cross reference stream object stream or failed to find xref table
        If Not IsNumeric(xrefType) Then Exit Function ' didn't find one, nothing to do
        Dim xrefStream As pdfValue
        Set xrefStream = GetValue(content, offset)
        If xrefStream.valueType <> PDF_ValueType.PDF_Object Then Exit Function ' wrong value returned, didn't find xref table
        Dim objStream As pdfStream
        Set objStream = xrefStream.Value.Value
        If Not objStream.Meta.Exists("/Type") Then Exit Function
        Dim obj As pdfValue
        Set obj = objStream.Meta.Item("/Type")
        If obj.valueType <> PDF_ValueType.PDF_Name Then Exit Function
        If Not IsMatch(obj.Value, "/XRef") Then Exit Function
        
        ' TODO merge? or replacing from stream if both exist
        Set trailer = New pdfValue
        Set trailer = pdfValueObj(pdfValueObj(objStream.Meta, "/Dictionary"), "/Trailer")
        
        ' get Size value from trailer (really stream dictionary), # of expected total entries
        Size = GetXrefSize(trailer)
        
        ' see how many entries in this xref stream
        Dim subSections As Collection
        If objStream.Meta.Exists("/Index") Then
            Set subSections = objStream.Meta.Item("/Index").Value
        Else ' use defaults
            Set subSections = New Collection
            subSections.Add pdfValueObj(0)
            subSections.Add pdfValueObj(Size)
        End If
        
        ' get the /W idth array for each entry
        If Not objStream.Meta.Exists("/W") Then Exit Function
        Dim widths(0 To 2) As Long
        Dim w As pdfValue
        Set w = objStream.Meta.Item("/W")
        widths(0) = CLng(w.Value.Item(1).Value)
        widths(1) = CLng(w.Value.Item(2).Value)
        widths(2) = CLng(w.Value.Item(3).Value)
        Set w = Nothing
        
        ' we need the uncompressed (un-/Filter'd) data
        Dim rawData() As Byte
        rawData = objStream.udata
        
        Dim objOffset As Long
        objOffset = 0
        Dim ndx As Long
        Dim idObj As pdfValue
        For ndx = 0 To subSections.Count - 1 Step 2
            ' get starting index (default value) and count of entries
            Set idObj = subSections.Item(ndx + 1)
            id = idObj.Value
            Set idObj = subSections.Item(ndx + 2)
            Count = idObj.Value
            Set idObj = Nothing
            
            ' read in object id and offset information
            Dim i As Long
            For i = 0 To Count - 1
                Dim recType As Long, field2 As Long, field3 As Long
                recType = getInt(rawData, objOffset, widths(0), 1) ' record type, 0=free,1=basic,2=embedded in obj stream
                field2 = getInt(rawData, objOffset, widths(1), 0)
                field3 = getInt(rawData, objOffset, widths(2), 0)
                
                Set entry = New xrefEntry
                entry.id = id
                id = id + 1
                Select Case recType
                    Case 0
                        entry.offset = 0
                        entry.nextFreeId = field2
                        entry.generation = field3
                        entry.isFree = True
                    Case 1
                        entry.offset = field2
                        entry.nextFreeId = 0
                        entry.generation = field3
                        entry.isFree = False
                    Case 2
                        entry.offset = field3
                        entry.nextFreeId = 0
                        entry.generation = 0
                        entry.isFree = False
                        entry.isEmbeded = True
                        entry.embedObjId = field2
                End Select
            
                ' add/replace to our catalog
                If xrefTable.Exists(entry.id) Then
                    ' Note: this is fine if we are reading in previous cross reference table,
                    ' it just means this object was updated (replaced) in the pdf
                    If entry.id <> 0 Then
                        Debug.Print "Warning: duplicate obj " & entry.id & " found!"
                    End If
                Else
                    xrefTable.Add entry.id, entry
                    Set entry = Nothing
                End If
            Next i
        Next ndx
    End If

    ' see if cross reference table has more parts
    Dim prevXref As pdfValue
    Set prevXref = GetDictionaryValue(trailer.Value, "/Prev")
    If prevXref.valueType <> PDF_ValueType.PDF_Null Then
        offset = prevXref.Value
        ParseXrefTable content, offset, trailer, xrefTable
    End If

    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function

' extract xref table (size # of xref entries)
' Note: trailer may be PDF_Null if cross reference stream, if so updated to xref stream dictionary << >>
Function GetXrefTable(ByRef content() As Byte, ByRef trailer As pdfValue) As Dictionary ' of xrefEntry
    On Error GoTo errHandler
        
    ' get offset of xref
    Dim offset As Long
    offset = GetXrefOffset(content)
    If offset <= 0 Then Exit Function ' didn't find one, nothing to do
    
    ' parse it
    Set GetXrefTable = ParseXrefTable(content, offset, trailer)
    
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function


' extracts/parses pdf object from raw pdf content()
' due to potential slowness uncompressing in VBA, stream object streams should be cached
' the sosCache is only used for stream object streams and only if provided
Function getObject(ByRef content() As Byte, ByRef xrefTable As Dictionary, ByVal Index As Long, ByRef sosCache As Dictionary) As pdfValue
    On Error GoTo errHandler
    Dim obj As pdfValue
    If xrefTable.Exists(Index) Then
        Dim entry As xrefEntry
        Set entry = xrefTable.Item(Index)
        If entry.isFree Or ((Not entry.isEmbeded) And (entry.offset <= 0)) Then GoTo nullValue
        Dim offset As Long: offset = entry.offset
        If offset > UBound(content) Then GoTo nullValue
        If entry.isEmbeded Then
            Dim cntrObjEntry As xrefEntry
            Set cntrObjEntry = xrefTable.Item(entry.embedObjId)
            Dim cntrObj As pdfValue
            ' try loading containing object (stream object stream) from cache before potentially uncompressing
            If Not sosCache Is Nothing Then
                If sosCache.Exists(entry.embedObjId) Then Set cntrObj = sosCache(entry.embedObjId)
            End If
            If cntrObj Is Nothing Then  ' not in cache or no cache provided
                Set cntrObj = getObject(content, xrefTable, entry.embedObjId, sosCache)
                If Not sosCache Is Nothing Then Set sosCache(entry.embedObjId) = cntrObj    ' add/update cache
            End If
            
            ' extract our embedded object
            If cntrObj.Value.valueType <> PDF_ValueType.PDF_Stream Then
                Debug.Print "Error! expecting stream object stream!"
                Stop
                GoTo nullValue
            End If
            Dim streamObjectStream As pdfStream
            Set streamObjectStream = cntrObj.Value.Value
            Dim buffer() As Byte
            buffer = streamObjectStream.udata
            If (UBound(buffer) - LBound(buffer)) > 0 Then
                ' parse embedded object data
                ' buffer has N sets of obj id# <whitespace> offset
                ' immediately followed by objects' data, note: /First
                ' should be used to determine where data starts when reading
                ' as additional data could exists between index and data
                ' we store as key=value, the object id=offset in embCatalog
                Dim embXRefTable As Dictionary: Set embXRefTable = New Dictionary
                Dim embOffset As Long
                Dim i As Long
                Dim firstOffset As Long
                Dim dict As Dictionary
                Set dict = streamObjectStream.Meta
                If dict.Exists("/First") Then
                    firstOffset = CLng(dict.Item("/First").Value)
                Else
                    firstOffset = -1
                End If
                Dim N As Long, id As Long, objOffset As Long
                Dim v As pdfValue
                If dict.Exists("/N") Then
                    N = CLng(dict.Item("/N").Value)
                    For i = 0 To N - 1
                        Set v = GetValue(buffer, embOffset)
                        id = v.Value
                        Set v = GetValue(buffer, embOffset)
                        objOffset = v.Value
                        embXRefTable.Add id, objOffset
                    Next i
                Else
                    Debug.Print "Missing count of embedded objects!"
                    If firstOffset >= 0 Then
                        Do
                            Set v = GetValue(buffer, embOffset)
                            id = v.Value
                            Set v = GetValue(buffer, embOffset)
                            offset = v.Value
                            embXRefTable.Add id, offset
                        Loop While embOffset < firstOffset
                    Else
                        Debug.Print "Error: unable to parse embedded object!"
                        Stop
                    End If
                End If
                ' curiousity check, spec says use /First
                SkipWhiteSpace buffer, embOffset
                If firstOffset < embOffset Then
                    Debug.Print "Warning: mini-catalog overlaps initial embedded object"
                ElseIf firstOffset > embOffset Then
                    Debug.Print "Warning: embedded object does not begin immediate after mini-catalog"
                End If
                embOffset = embXRefTable.Item(entry.id) + firstOffset
                Set obj = New pdfValue
                obj.id = entry.id
                obj.valueType = PDF_ValueType.PDF_Object
                Set obj.Value = GetValue(buffer, embOffset)
            Else
                Debug.Print "Error reading embedded object!"
                Stop
            End If
        Else
            Set obj = GetValue(content, offset)
        End If
    Else
nullValue:
        Set obj = New pdfValue
        obj.valueType = PDF_ValueType.PDF_Null
    End If
    
    Set getObject = obj
    Set obj = Nothing
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function


Function GetRootObject(ByRef content() As Byte, ByRef trailer As pdfValue, ByRef xrefTable As Dictionary) As pdfValue
    On Error GoTo errHandler
    Dim offset As Long
    Dim root As pdfValue
    ' get either reference or /Root object itself
    Set root = GetRoot(trailer)
    If root.valueType = PDF_ValueType.PDF_Reference Then
        Set root = getObject(content, xrefTable, root.Value, Nothing)
    'ElseIf root.valueType = PDF_ValueType.PDF_Object Then
    End If
    
    ' returns PDF_Dictionary object or PDF_Null
    Set GetRootObject = root
    Set root = Nothing
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Function


Function GetInfoObject(ByRef content() As Byte, ByRef trailer As pdfValue, ByRef xrefTable As Dictionary) As pdfValue
    On Error GoTo errHandler
    Dim Info As pdfValue
    ' get either reference or /Info object itself
    Set Info = GetInfo(trailer)
    If Info.valueType = PDF_ValueType.PDF_Reference Then
        Set Info = getObject(content, xrefTable, Info.Value, Nothing)
    'ElseIf info.valueType = PDF_ValueType.PDF_Object Then
    End If
    
    ' returns PDF_Dictionary object or PDF_Null
    Set GetInfoObject = Info
    Set Info = Nothing
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Function


' updates objects Dictionary with all objects under root node, indexed by object id, i.e. loads a chunk of the PDF document
Sub GetObjectsInTree(ByRef root As pdfValue, ByRef content() As Byte, ByRef xrefTable As Dictionary, ByRef objects As Dictionary, ByRef sosCache As Dictionary)
    On Error GoTo errHandler
    Dim obj As pdfValue
    Dim v As Variant
    DoEvents
    'Debug.Print BytesToString(serialize(root))
    Select Case root.valueType
        Case PDF_ValueType.PDF_Boolean, PDF_ValueType.PDF_Comment, PDF_ValueType.PDF_Integer, PDF_ValueType.PDF_Name, PDF_ValueType.PDF_Null, PDF_ValueType.PDF_Real, PDF_ValueType.PDF_String
            ' Nothing to do
        Case PDF_ValueType.PDF_Array
            For Each v In root.Value
                Set obj = v
                GetObjectsInTree obj, content, xrefTable, objects, sosCache
            Next v
        Case PDF_ValueType.PDF_Dictionary
            For Each v In root.Value.Items
                Set obj = v
                GetObjectsInTree obj, content, xrefTable, objects, sosCache
            Next v
        Case PDF_ValueType.PDF_Object
            GetObjectsInTree root.Value, content, xrefTable, objects, sosCache
        Case PDF_ValueType.PDF_Reference
            ' we need to load object
            If Not objects.Exists(CLng(root.Value)) Then
                Set obj = getObject(content, xrefTable, root.Value, sosCache)
                objects.Add CLng(root.Value), obj
                GetObjectsInTree obj, content, xrefTable, objects, sosCache
            End If
        Case PDF_ValueType.PDF_Stream
            Dim stream As pdfStream
            Set stream = root.Value
            GetObjectsInTree stream.stream_meta, content, xrefTable, objects, sosCache
            GetObjectsInTree stream.stream_data, content, xrefTable, objects, sosCache
        Case PDF_ValueType.PDF_StreamData
            ' Nothing to do
        Case PDF_ValueType.PDF_Trailer
            Stop ' ???
        Case Else
    End Select
    Exit Sub
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Sub

