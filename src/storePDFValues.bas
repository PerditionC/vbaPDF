Attribute VB_Name = "storePDFValues"
' converts a PDF Value into a Byte() array
Option Explicit


Function serialize(ByRef Value As pdfValue, Optional ByVal baseId As Long = 0) As Byte()
    On Error GoTo errHandler
    Dim objStr As String: objStr = vbNullString
    Dim objBytes() As Byte
    Dim IsBytes As Boolean ' for most we convert at end, but stream we leave as Byte()
    Dim v As Variant
    Dim pv As pdfValue
    Dim firstPass As Boolean
    
    Select Case Value.valueType
        Case PDF_ValueType.PDF_Null
            objStr = "null"
        Case PDF_ValueType.PDF_Name
            objStr = Value.Value
        Case PDF_ValueType.PDF_Boolean
            If Value.Value Then
                objStr = "true"
            Else
                objStr = "false"
            End If
        Case PDF_ValueType.PDF_Integer
            objStr = Format(CLng(Value.Value), "0")
        Case PDF_ValueType.PDF_Real
            objStr = CDbl(Value.Value) ' dont' format as we want all current digits stored, it won't add extra 0s anyway
            ' ensure has .0 if whole number
            If InStr(1, objStr, ".", vbBinaryCompare) < 1 Then objStr = objStr & ".0"
        Case PDF_ValueType.PDF_String ' TODO escape )
            objStr = "(" & Value.Value & ")"
        Case PDF_ValueType.PDF_Array
            objStr = "[ "
            firstPass = True
            For Each v In Value.Value
                Set pv = v
                If Not firstPass Then objStr = objStr & " "
                firstPass = False
                objStr = objStr & BytesToString(serialize(pv, baseId))
            Next v
            objStr = objStr & " ]"
        Case PDF_ValueType.PDF_Dictionary
            objStr = "<<" & vbLf
            Dim dict As Dictionary
            Set dict = Value.Value
            firstPass = True
            For Each v In dict.Keys
                Set pv = dict.Item(v)
                If Not firstPass Then objStr = objStr & vbLf
                firstPass = False
                objStr = objStr & CStr(v) & " "
                objStr = objStr & BytesToString(serialize(pv, baseId))
            Next v
            If Right(objStr, 1) <> vbLf Then objStr = objStr & vbLf
            objStr = objStr & ">>" & vbLf
        Case PDF_ValueType.PDF_Stream       ' actual stream object with dictionary and data
            Dim stream As pdfStream
            Set stream = Value.Value
            IsBytes = True
            objBytes = serialize(stream.stream_meta, baseId)
            CopyBytes serialize(stream.stream_data, baseId), objBytes, 0, UBound(objBytes) + 1
        Case PDF_ValueType.PDF_StreamData   ' represents only stream ... endstream portion
            IsBytes = True
            objBytes = StringToBytes("stream" & vbLf)
            CopyBytes Value.Value, objBytes, 0, UBound(objBytes) + 1
            CopyBytes StringToBytes(vbLf & "endstream" & vbLf), objBytes, 0, UBound(objBytes) + 1
        
    ' to simplify processing, not one of 9 basic types either
        Case PDF_ValueType.PDF_Object       ' id generation obj << dictionary >> endobj
            IsBytes = True
            objBytes = StringToBytes(baseId + Value.id & " " & Value.generation & " obj" & vbLf)
            CopyBytes serialize(Value.Value, baseId), objBytes, 0, UBound(objBytes) + 1
            If objBytes(UBound(objBytes)) <> 10 Then CopyBytes StringToBytes(vbLf), objBytes, 0, UBound(objBytes) + 1
            CopyBytes StringToBytes("endobj" & vbLf), objBytes, 0, UBound(objBytes) + 1
        Case PDF_ValueType.PDF_Reference    ' indirect object
            ' Note: if indirect reference to /Parent and that obj is not in current set, will have wrong id, offset correct id by -baseId in Reference object prior to saving
            objStr = baseId + Value.Value & " " & Value.generation & " R"
        Case PDF_ValueType.PDF_Comment
            objStr = Value.Value & vbLf
        Case PDF_ValueType.PDF_Trailer
            objStr = "trailer" & vbLf
            objStr = objStr & BytesToString(serialize(Value.Value, baseId))
        Case Else
            Stop ' ???
    End Select
    
    If IsBytes Then
        serialize = objBytes
    Else
        serialize = StringToBytes(objStr)
    End If
    
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function

