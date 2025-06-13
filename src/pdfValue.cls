VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pdfValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' stores a PDF value type
Option Explicit

Public id As Long
Public generation As Long
Public valueType As PDF_ValueType
Public Value As Variant


' is this merely an object containing a pdf dictionary? i.e. is Value a pdfValue of PDF_Dictionary
Function isDictionaryObject() As Boolean
    If (valueType = PDF_ValueType.PDF_Object) And (Not Value Is Nothing) Then
        Dim pdfDict As pdfValue: Set pdfDict = Value
        isDictionaryObject = (pdfDict.valueType = PDF_ValueType.PDF_Dictionary)
    End If
End Function

' returns vba Dictionary instance contained in pdf dictionary
' Warning: returns Nothing if object lacks dictionary to return
Function asDictionary() As Dictionary
    On Error GoTo errHandler
    If ((valueType = PDF_ValueType.PDF_Object) Or (valueType = PDF_ValueType.PDF_Dictionary) Or (valueType = PDF_ValueType.PDF_Trailer)) _
    And (Not IsEmpty(Value)) Then
        Dim pdfDict As pdfValue
        If (valueType = PDF_ValueType.PDF_Dictionary) Then
            Set pdfDict = Me
        Else
            Set pdfDict = Value
        End If
        Set asDictionary = pdfDict.Value
    End If
    Exit Function
errHandler:
    Debug.Print Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function


' Helper function to convert pdf string to byte array (0..N-1)
' escaping values as needed
Private Function pdfStringToBytes(ByRef str As String, Optional ByVal asHex As Boolean = False) As Byte()
    On Error GoTo errHandler
    Dim strLen As Long
    Dim ucBytes() As Byte, bytes() As Byte
    strLen = Len(str)
    If strLen < 1 Then
        pdfStringToBytes = StringToBytes("()") ' empty PDF string
        Exit Function
    End If
    
    ' we need to check if any Unicode values and if so we encode in UTF-8
    ' to avoid more expensive check if >255 value used, we simply always encode as UTF-8 unless ASCII (all values < 128)
    ucBytes = StringToUtf8Bytes(str, BOM:=False)  ' Note: if only ASCII then equivalent to UTF-8 byte array
    If UBound(ucBytes) > strLen Then ucBytes = AddUtf8BOM(ucBytes)    ' something >128 and had to be expanded to multiple bytes
    
    Dim ndx As Long
    ' now determine if we are returning bytes as (string) or <## ## ##> format
    If asHex Then
        ' we need space for <## > assume 3 bytes per byte + 2 for enclosing <>
        ReDim bytes(0 To ((UBound(ucBytes) + 1) * 3) + 1)
        bytes(0) = Asc("<")
        bytes(UBound(bytes)) = Asc(">")
        For ndx = 0 To UBound(ucBytes)
            Dim hexValue As String
            hexValue = Hex(ucBytes(ndx))
            If Len(hexValue) > 1 Then
                bytes((ndx * 3) + 1) = Asc(Left(hexValue, 1))
            Else
                bytes((ndx * 3) + 1) = &H30 ' Asc("0")
            End If
            bytes((ndx * 3) + 2) = Asc(Right(hexValue, 1))
            bytes((ndx * 3) + 3) = &H20 ' Asc(" ")
        Next ndx
    Else
        ReDim bytes(0 To UBound(ucBytes) + 2)
        bytes(0) = Asc("(")
        bytes(UBound(bytes)) = Asc(")")
        'CopyBytes ucBytes, bytes, 0, 1
        ' we need to escape ) and \ characters (but only unbalanced ) actually need escaping)
        Dim byteNdx As Long: byteNdx = 1
        For ndx = 0 To UBound(ucBytes)
            ' for now slow, we resize on every escaping
            If (ucBytes(ndx) = 41) Or (ucBytes(ndx) = 92) Then  ' Asc(")") or Asc("\")
                ReDim Preserve bytes(0 To UBound(bytes) + 1)
                bytes(byteNdx) = 92 ' Asc("\")
                byteNdx = byteNdx + 1
            End If
            ' copy over byte, may be preceed by \ if needed
            bytes(byteNdx) = ucBytes(ndx)
            byteNdx = byteNdx + 1
        Next ndx
    End If
    
    pdfStringToBytes = bytes
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function


' convert pdf Value (Me) into a Byte() array as stored in pdf file
Function serialize(Optional ByVal baseId As Long = 0) As Byte()
    On Error GoTo errHandler
    Dim Value As pdfValue: Set Value = Me
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
        Case PDF_ValueType.PDF_String
            ' convert to bytes and escape values as needed
            IsBytes = True
            objBytes = pdfStringToBytes(Value.Value)
        Case PDF_ValueType.PDF_Array
            objStr = "[ "
            firstPass = True
            For Each v In Value.Value
                Set pv = v
                If Not firstPass Then objStr = objStr & " "
                firstPass = False
                objStr = objStr & BytesToString(pv.serialize(baseId))
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
                objStr = objStr & BytesToString(pv.serialize(baseId))
            Next v
            If Right(objStr, 1) <> vbLf Then objStr = objStr & vbLf
            objStr = objStr & ">>" & vbLf
        Case PDF_ValueType.PDF_Stream       ' actual stream object with dictionary and data
            Dim stream As pdfStream
            Set stream = Value.Value
            IsBytes = True
            objBytes = stream.stream_meta.serialize(baseId)
            CopyBytes stream.stream_data.serialize(baseId), objBytes, 0, UBound(objBytes) + 1
        Case PDF_ValueType.PDF_StreamData   ' represents only stream ... endstream portion
            IsBytes = True
            objBytes = StringToBytes("stream" & vbLf)
            CopyBytes Value.Value, objBytes, 0, UBound(objBytes) + 1
            CopyBytes StringToBytes(vbLf & "endstream" & vbLf), objBytes, 0, UBound(objBytes) + 1
        
    ' to simplify processing, not one of 9 basic types either
        Case PDF_ValueType.PDF_Object       ' id generation obj << dictionary >> endobj
            IsBytes = True
            objBytes = StringToBytes(baseId + Value.id & " " & Value.generation & " obj" & vbLf)
            Dim pdfObj As pdfValue: Set pdfObj = Value.Value
            CopyBytes pdfObj.serialize(baseId), objBytes, 0, UBound(objBytes) + 1
            If objBytes(UBound(objBytes)) <> 10 Then CopyBytes StringToBytes(vbLf), objBytes, 0, UBound(objBytes) + 1
            CopyBytes StringToBytes("endobj" & vbLf), objBytes, 0, UBound(objBytes) + 1
        Case PDF_ValueType.PDF_Reference    ' indirect object
            ' Note: if indirect reference to /Parent and that obj is not in current set, will have wrong id, offset correct id by -baseId in Reference object prior to saving
            objStr = baseId + Value.Value & " " & Value.generation & " R"
        Case PDF_ValueType.PDF_Comment
            objStr = Value.Value & vbLf
        Case PDF_ValueType.PDF_Trailer
            Dim pdfTrailer As pdfValue: Set pdfTrailer = Value.Value
            objStr = "trailer" & vbLf
            objStr = objStr & BytesToString(pdfTrailer.serialize(baseId))
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

