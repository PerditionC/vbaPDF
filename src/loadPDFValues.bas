Attribute VB_Name = "loadPDFValues"
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
    PDF_Comment
    PDF_Trailer
    
    ' markers, no actual value returned
    PDF_EndOfArray
    PDF_EndOfDictionary
    PDF_EndOfStream
    PDF_EndOfObject
End Enum

Function GetValueType(ByRef bytes() As Byte, ByVal offset As Long) As PDF_ValueType
    On Error GoTo errHandler
    GetValueType = PDF_ValueType.PDF_Null
    
    Dim token As String: token = Chr(bytes(offset))
    Do While IsWhiteSpace(token)
        offset = offset + 1
        If offset > UBound(bytes) Then Exit Function    ' return null type if end of data
        token = Chr(bytes(offset))
    Loop
    
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
        Case "<"
            If IsMatch("<", Chr(bytes(offset + 1))) Then
                GetValueType = PDF_ValueType.PDF_Dictionary ' <<...>>
            Else
                GetValueType = PDF_ValueType.PDF_String ' hex string <####>
            End If
        Case "%"
            GetValueType = PDF_ValueType.PDF_Comment
            
        Case "e", ">", "]"
            tmpStr = GetWord(bytes, offset)
            Select Case LCase(tmpStr)
                Case "]"
                    GetValueType = PDF_ValueType.PDF_EndOfArray
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
        ProcessName = left(name, ndx - 1)
        s = Mid(name, ndx + 1, 2) ' get hex digits
        N = CLng("&H" & s) ' convert from hex to Long
        s = ProcessName & Chr(N) & Mid(name, ndx + 3) ' combine with hex digit as character
        ProcessName = ProcessName(s) ' for now just recursively call to handle multiple encoded values, really should be a loop!
    Else
        ProcessName = name
    End If
End Function


' returns a value loaded for a PDF
' updates offset to next non-whitespace byte after this value is loaded
' Note: meta is only used for stream object
Function GetValue(ByRef bytes() As Byte, ByRef offset As Long, Optional Meta As Dictionary = Nothing) As pdfValue
    On Error GoTo errHandler
    DoEvents
    Dim result As pdfValue: Set result = New pdfValue
    result.id = 0
    result.generation = 0
    result.Value = Empty
    result.valueType = GetValueType(bytes, offset)
    Set GetValue = result
    
    If offset > UBound(bytes) Then Exit Function        ' return null type if end of daa
    Dim token As String: token = Chr(bytes(offset))
    Do While IsWhiteSpace(token)
        offset = offset + 1
        If offset > UBound(bytes) Then Exit Function    ' return null type if end of data
        token = Chr(bytes(offset))
    Loop
    
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
            tmpStr = vbNullString
            If bytes(offset) = Asc("(") Then
                offset = offset + 1
                Do While bytes(offset) <> Asc(")")
                    If bytes(offset) = Asc("\") Then ' escaped value
                        offset = offset + 1
                        Select Case Chr(bytes(offset))
                            Case "n"
                                tmpStr = tmpStr & vbLf
                            Case "r"
                                tmpStr = tmpStr & vbCr
                            Case "t"
                                tmpStr = tmpStr & vbTab
                            Case "f"
                                tmpStr = tmpStr & Chr(&HC)  ' formfeed
                            Case "b"
                                tmpStr = tmpStr & Chr(&H8)  ' backspace
                            Case "\"
                                tmpStr = tmpStr & "\"
                            Case ")"
                                tmpStr = tmpStr & ")"
                            Case "("
                                tmpStr = tmpStr & "("
                            Case Else
                                ' TODO
                                Stop
                        End Select
                    Else
                        tmpStr = tmpStr & Chr(bytes(offset))
                    End If
                    DoEvents
                    offset = offset + 1
                Loop
                offset = offset + 1 ' skip past ending ")"
            Else
                offset = offset + 1
                Do While bytes(offset) <> 62 'Asc(">")
                    ' get 2 hex digits, ignoring whitespace, may end with odd # of hex digits
                    While IsWhiteSpace(bytes(offset))
                        offset = offset + 1
                        DoEvents
                    Wend
                    Dim HexStr As String
                    HexStr = Chr(bytes(offset))
                    offset = offset + 1
                    While IsWhiteSpace(bytes(offset))
                        offset = offset + 1
                        DoEvents
                    Wend
                    If bytes(offset) <> 62 Then ' Asc(">")
                        HexStr = HexStr & Chr(bytes(offset))
                        offset = offset + 1
                    Else
                        HexStr = HexStr & "0"
                    End If
                    
                    tmpStr = tmpStr & Chr(CLng("&H" & HexStr))
                    DoEvents
                Loop
                offset = offset + 1 ' skip past ending ">"
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
                Do While Not IsMatch(left(tmpStr, 9), "endstream")
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
                If Right(dataStr, 1) = vbLf Then dataStr = left(dataStr, Len(dataStr) - 1)
                If Right(dataStr, 1) = vbCr Then dataStr = left(dataStr, Len(dataStr) - 1)
                dataBytes = StringToBytes(dataStr)
            End If
            result.Value = dataBytes
            If Not IsMatch(left(tmpStr, 9), "endstream") Then Stop ' error unexpected token found!
            
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
End Function


Function GetXrefOffset(ByRef content() As Byte) As Long
    On Error GoTo errHandler
    Dim offset As Long
    offset = FindToken(content, "startxref", searchBackward:=True)
    offset = SkipWhiteSpace(content, offset + Len("startxref"))
    GetXrefOffset = CLng(GetWord(content, offset))
    ' should immediately be followed by end of file marker
    Dim Value As pdfValue: Set Value = GetValue(content, offset)
    If Value.valueType <> PDF_ValueType.PDF_Comment Then Stop ' error, expected "%%EOF"
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


' given a pdfValue dictionary object, returns the values associated with a given name
Function GetDictionaryValue(ByRef Value As pdfValue, ByVal name As String) As pdfValue
    On Error GoTo errHandler
    If Value.Value.Exists(name) Then
        Set GetDictionaryValue = Value.Value.Item(name)
    Else
        Set GetDictionaryValue = New pdfValue
        GetDictionaryValue.valueType = PDF_ValueType.PDF_Null
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
Function ParseXrefTable(ByRef content() As Byte, ByRef offset As Long, ByRef trailer As pdfValue, Optional ByRef catalog As Dictionary = Nothing) As Dictionary ' of xrefEntry
    Dim primaryXref As Boolean ' = False
    If catalog Is Nothing Then
        Set catalog = New Dictionary
        primaryXref = True
    End If
    Set ParseXrefTable = catalog

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
                If catalog.Exists(entry.id) Then
                    If (entry.id <> 0) And primaryXref Then
                        Debug.Print "Warning: duplicate obj " & entry.id & " found!"
                        Stop
                    End If
                Else
                    catalog.Add entry.id, entry
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
            Debug.Print "Warning: PDF did not find /Size catalog entries - missing " & Size & " entries."
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
        If objStream.Meta.Exists("/Filter") Then
            ' TODO support all /Filter types
            Dim filter As String
            filter = objStream.Meta.Item("/Filter").Value
            Select Case LCase(filter)
                Case "/flatedecode"
                    Dim sourceLen As Long
                    sourceLen = UBound(objStream.data)
                    ReDim rawData(0 To (sourceLen * 4))
                    Dim rawLength As Long ' as count of bytes
                    'zlib_Inflate.uncompress rawData, destLen, objStream.data, sourceLen
                    'libdeflate_inflate objStream.data, 2, rawData, destLen
                    
                    sourceLen = 2   ' skip past zlib wrapper to raw Deflate data
                    If Not inflate2(objStream.data, rawData, sourceLen, rawLength) Then
                        Debug.Print "Error decompressing!"
                        Stop
                    End If
                    
                    ' do we need to decode uncompressed stream?
                    Dim predictor As Long, columns As Long
                    ' set default values if not otherwise specified
                    predictor = 1: columns = 1
                    If objStream.Meta.Exists("/DecodeParms") Then
                        Dim pdfV As pdfValue
                        Set pdfV = objStream.Meta.Item("/DecodeParms")
                        Dim decodeParms As Dictionary
                        Set decodeParms = pdfV.Value
                        If decodeParms.Exists("/Predictor") Then
                            Set pdfV = decodeParms.Item("/Predictor")
                            predictor = pdfV.Value
                        End If
                        ' should only be supplied if predictor > 1, but we can load value regardless, only used if predictor > 1
                        If decodeParms.Exists("/Columns") Then
                            Set pdfV = decodeParms.Item("/Columns")
                            columns = pdfV.Value
                        End If
                        Set pdfV = Nothing
                    End If
                    ' if predictor > 1 then we need to reverse differencing done prior to compression/encoding
                    Dim rowOffset As Long, rowIndex As Long
                    Select Case predictor
                        Case 1 ' default no prediction
                            ' nothing to do
                        Case 2 ' TIFF predictor
                            Stop ' not implemented
                        
                        ' Note: regardless of predictor specified, if PNG filter then each row should have a predicator tag, need not match predictor value
                        ' 10 through 15 are defined in PNG RFC 2083 specification, see Chapter 6, Filters
                        Case 10, 11, 12, 13, 14, 15
                            ' actual data is smaller, we need to buffer our data
                            Dim buffer() As Byte
                            ReDim buffer(0 To UBound(rawData))
                            Dim bufferIndex As Long
                            ' loop through all the data, assuming each row is columns bytes wide
                            ' note: each row has columns bytes of data + 1 for PNG filter type, except 1st row
                            columns = columns + 1
                            Dim rowPredictor As Long
                            For rowIndex = LBound(rawData) To UBound(rawData)
                                'rowOffset = LBound(rawData) + ((rowIndex - LBound(rawData)) Mod columns)
                                rowOffset = rowIndex Mod columns ' simplify since LBound(rawData) = 0
                                ' get this rows predicator tag, Note: 1st row doesn't have tag byte
                                If rowOffset = 0 Then ' this byte is a tag byte for new row
                                    rowPredictor = rawData(rowIndex)
                                    ' advance to actual data
                                    rowIndex = rowIndex + 1
                                    If rowIndex > UBound(rawData) Then Exit For
                                End If
                                                        
                                Select Case rowPredictor
                                    Case 0, 10 ' PNG, None on all rows
                                        buffer(bufferIndex) = rawData(rowIndex)
                                    Case 1, 11 ' PNG, Sub on all rows
                                        Stop ' no implemented
                                    Case 2, 12 ' PNG, Up on all rows
                                        If rowIndex >= columns Then ' assume 1st row, with prior row values always 0, so no change to values
                                            ' assume rawData(rowIndex - columns) is the data byte 1 row Up, add that to current difference value at rowIndex
                                            'buffer(bufferIndex) = (CLng(rawData(rowIndex - columns)) + CLng(rawData(rowIndex))) Mod 256
                                            'Debug.Print Hex(buffer(bufferIndex)) & " - ";
                                            buffer(bufferIndex) = (CLng(buffer(bufferIndex - (columns - 1))) + CLng(rawData(rowIndex))) Mod 256
                                            'Debug.Print Hex(buffer(bufferIndex))
                                        Else
                                            buffer(bufferIndex) = rawData(rowIndex)
                                        End If
                                    Case 3, 13 ' PNG, Average on all rows
                                        Stop ' no implemented
                                    Case 4, 14 ' PNG, Paeth on all rows
                                        Stop ' no implemented
                                    Case 15 ' PNG, Optimal (per row determination)
                                        Stop ' no implemented, error not valid PNG filter value
                                            
                                    Case Else
                                        Debug.Print "Error: unsupported or invalid predictor found - " & predictor
                                        Stop
                                End Select ' rowPredictor
                                
                                bufferIndex = bufferIndex + 1
                                DoEvents
                            Next rowIndex
                            ' swap out with our smaller buffer omitting predicator data
                            rawData = buffer
                            ReDim Preserve rawData(0 To (bufferIndex - 1))
                            
                        Case Else
                            Debug.Print "Error: unsupported or invalid predictor found - " & predictor
                            Stop
                    End Select ' predictor
                Case "/lzwdecode"
                    Stop ' TODO
                Case Else
                    Stop ' not yet supported!
            End Select
        Else
            rawData = objStream.data
        End If
        
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
                If catalog.Exists(entry.id) Then
                    If entry.id <> 0 Then
                        Debug.Print "Warning: duplicate obj " & entry.id & " found!"
                        Stop
                    End If
                Else
                    catalog.Add entry.id, entry
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
        ParseXrefTable content, offset, trailer, catalog
    End If
End Function

' extract xref table (size # of xref entries)
' Note: trailer may be PDF_Null if cross reference stream, if so updated to xref stream dictionary << >>
Function GetXrefTable(ByRef content() As Byte, ByRef trailer As pdfValue) As Dictionary ' of xrefEntry
    On Error GoTo errHandler
    Dim catalog As Dictionary
        
    ' get offset of xref
    Dim offset As Long
    offset = GetXrefOffset(content)
    If offset <= 0 Then Exit Function ' didn't find one, nothing to do
    
    ' parse it
    Set catalog = ParseXrefTable(content, offset, trailer)
    
    Set GetXrefTable = catalog
    Set catalog = Nothing
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function


Function GetObject(ByRef content() As Byte, ByRef catalog As Dictionary, ByVal Index As Long) As pdfValue
    On Error GoTo errHandler
    Dim obj As pdfValue
    If catalog.Exists(Index) Then
        Dim entry As xrefEntry
        Set entry = catalog.Item(Index)
        If entry.isFree Or ((Not entry.isEmbeded) And (entry.offset <= 0)) Then GoTo nullValue
        Dim offset As Long: offset = entry.offset
        If offset > UBound(content) Then GoTo nullValue
        If entry.isEmbeded Then
            Dim cntrObjEntry As xrefEntry
            Set cntrObjEntry = catalog.Item(entry.embedObjId)
            Dim cntrObj
            Set cntrObj = GetObject(content, catalog, entry.embedObjId)
            
            ' extract our embedded object
            Dim bufSize As Long
            If cntrObj.Value.Value.Meta.Exists("/Length") Then
                Dim dict As Dictionary
                Set dict = cntrObj.Value.Value.Meta
                bufSize = CLng(dict.Item("/Length").Value)
            Else
                bufSize = UBound(cntrObj.Value.Value.data) * 4
            End If
            Dim buffer() As Byte
            ReDim buffer(0 To bufSize - 1)
            ReDim buffer(0 To bufSize * 2)
            Dim inOff As Long, outSize As Long
            Dim cbuf() As Byte
            cbuf = cntrObj.Value.Value.data
            inOff = 2
            If inflate2(cbuf, buffer, inOff, outSize) Then
                ' parse embedded object data
                ' buffer has N sets of obj id# <whitespace> offset
                ' immediately followed by objects' data, note: /First
                ' should be used to determine where data starts when reading
                ' as additional data could exists between index and data
                ' we store as key=value, the object id=offset in embCatalog
                Dim embCatalog As Dictionary: Set embCatalog = New Dictionary
                Dim embOffset As Long
                Dim i As Long
                Dim firstOffset As Long
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
                        embCatalog.Add id, objOffset
                    Next i
                Else
                    Debug.Print "Missing count of embedded objects!"
                    If firstOffset >= 0 Then
                        Do
                            Set v = GetValue(buffer, embOffset)
                            id = v.Value
                            Set v = GetValue(buffer, embOffset)
                            offset = v.Value
                            embCatalog.Add id, offset
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
                embOffset = embCatalog.Item(entry.id) + firstOffset
                Set obj = New pdfValue
                obj.id = entry.id
                obj.valueType = PDF_ValueType.PDF_Object
                Set obj.Value = GetValue(buffer, embOffset)
            Else
                Debug.Print "Error inflating embedded object!"
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
    
    Set GetObject = obj
    Set obj = Nothing
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function


Function GetRootObject(ByRef content() As Byte, ByRef trailer As pdfValue, ByRef catalog As Dictionary) As pdfValue
    On Error GoTo errHandler
    Dim offset As Long
    Dim root As pdfValue
    ' get either reference or /Root object itself
    Set root = GetRoot(trailer)
    If root.valueType = PDF_ValueType.PDF_Reference Then
        Set root = GetObject(content, catalog, root.Value)
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


Function GetInfoObject(ByRef content() As Byte, ByRef trailer As pdfValue, ByRef catalog As Dictionary) As pdfValue
    On Error GoTo errHandler
    Dim offset As Long
    Dim info As pdfValue
    ' get either reference or /Info object itself
    Set info = GetInfo(trailer)
    If info.valueType = PDF_ValueType.PDF_Reference Then
        Set info = GetObject(content, catalog, info.Value)
    'ElseIf info.valueType = PDF_ValueType.PDF_Object Then
    End If
    
    ' returns PDF_Dictionary object or PDF_Null
    Set GetInfoObject = info
    Set info = Nothing
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Function


' updates objects Dictionary with all objects under root node, indexed by object id, i.e. loads a chunk of the PDF document
Sub GetObjectsInTree(ByRef root As pdfValue, ByRef content() As Byte, ByRef catalog As Dictionary, ByRef objects As Dictionary)
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
                GetObjectsInTree obj, content, catalog, objects
            Next v
        Case PDF_ValueType.PDF_Dictionary
            For Each v In root.Value.Items
                Set obj = v
                GetObjectsInTree obj, content, catalog, objects
            Next v
        Case PDF_ValueType.PDF_Object
            GetObjectsInTree root.Value, content, catalog, objects
        Case PDF_ValueType.PDF_Reference
            ' we need to load object
            If Not objects.Exists(CLng(root.Value)) Then
                Set obj = GetObject(content, catalog, root.Value)
                objects.Add CLng(root.Value), obj
                GetObjectsInTree obj, content, catalog, objects
            End If
        Case PDF_ValueType.PDF_Stream
            Dim stream As pdfStream
            Set stream = root.Value
            GetObjectsInTree stream.stream_meta, content, catalog, objects
            GetObjectsInTree stream.stream_data, content, catalog, objects
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

