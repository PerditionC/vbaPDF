VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pdfValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' stores a PDF value type
Option Explicit

Private m_id As Long
Private m_generation As Long
Public valueType As PDF_ValueType
Public value As Variant

' to help maintain id across indirect references, PDF_Object provides a PDF_Reference object to be used by others
Private m_referenceObj As pdfValue


' keep our referencObj in sync
Public Property Let id(ByVal newId As Long)
    m_id = newId
    If valueType = PDF_ValueType.PDF_Object Then
        If Not referenceObj Is Nothing Then
            referenceObj.value = m_id
        End If
    End If
End Property
Public Property Get id() As Long
    id = m_id
End Property

' keep our referencObj in sync
Public Property Let generation(ByVal newGeneration As Long)
    m_generation = newGeneration
    If valueType = PDF_ValueType.PDF_Object Then
        If Not m_referenceObj Is Nothing Then
            m_referenceObj.generation = m_generation
        End If
    End If
End Property
Public Property Get generation() As Long
    generation = m_generation
End Property


' returns an indirect reference object to self
Friend Property Set referenceObj(ByRef refObj As pdfValue)
    Set m_referenceObj = refObj
End Property
Public Property Get referenceObj() As pdfValue
    ' create if not already exists
    If m_referenceObj Is Nothing Then
        Set m_referenceObj = pdfValue.NewValue(Me, "/Reference")
    End If
    
    ' always return same object so all references updated if id changes
    Set referenceObj = m_referenceObj
End Property


#If False Then
' is this merely an object containing a pdf dictionary? i.e. is Value a pdfValue of PDF_Dictionary
Function isDictionaryObject() As Boolean
    If (valueType = PDF_ValueType.PDF_Object) And (Not value Is Nothing) Then
        Dim pdfDict As pdfValue: Set pdfDict = value
        isDictionaryObject = (pdfDict.valueType = PDF_ValueType.PDF_Dictionary)
    End If
End Function
#End If

' returns vba Dictionary instance contained in pdf dictionary
' Warning: returns Nothing if object lacks dictionary to return, hasKey()
Function asDictionary() As Dictionary
    On Error GoTo errHandler
    If ((valueType = PDF_ValueType.PDF_Object) Or (valueType = PDF_ValueType.PDF_Dictionary) Or (valueType = PDF_ValueType.PDF_Trailer)) _
    And (Not IsEmpty(value)) Then
        Dim pdfDict As pdfValue
        If (valueType = PDF_ValueType.PDF_Dictionary) Then
            Set pdfDict = Me
        Else
            Set pdfDict = value
        End If
        Set asDictionary = pdfDict.value
    End If
    Exit Function
errHandler:
    Debug.Print Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function

' returns True if Dictionary object contains key
Function hasKey(ByRef key As Variant) As Boolean
    On Error GoTo errHandler
    Dim dict As Dictionary
    Set dict = asDictionary()
    If Not dict Is Nothing Then hasKey = dict.Exists(key)
    'Else return False
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
Public Function serialize(Optional ByVal baseId As Long = 0) As Byte()
    On Error GoTo errHandler
    Dim value As pdfValue: Set value = Me
    Dim objStr As String: objStr = vbNullString
    Dim objBytes() As Byte
    Dim IsBytes As Boolean ' for most we convert at end, but stream we leave as Byte()
    Dim v As Variant
    Dim pv As pdfValue
    Dim firstPass As Boolean
    
    Select Case value.valueType
        Case PDF_ValueType.PDF_Null
            objStr = "null"
        Case PDF_ValueType.PDF_Name
            objStr = value.value
        Case PDF_ValueType.PDF_Boolean
            If value.value Then
                objStr = "true"
            Else
                objStr = "false"
            End If
        Case PDF_ValueType.PDF_Integer
            objStr = Format(CLng(value.value), "0")
        Case PDF_ValueType.PDF_Real
            objStr = CDbl(value.value) ' dont' format as we want all current digits stored, it won't add extra 0s anyway
            ' ensure has .0 if whole number
            If InStr(1, objStr, ".", vbBinaryCompare) < 1 Then objStr = objStr & ".0"
        Case PDF_ValueType.PDF_String
            ' convert to bytes and escape values as needed
            IsBytes = True
            objBytes = pdfStringToBytes(value.value)
        Case PDF_ValueType.PDF_Array
            objStr = "[ "
            firstPass = True
            For Each v In value.value
                Set pv = v
                If Not firstPass Then objStr = objStr & " "
                firstPass = False
                objStr = objStr & BytesToString(pv.serialize(baseId))
            Next v
            objStr = objStr & " ]"
        Case PDF_ValueType.PDF_Dictionary
            objStr = "<<" & vbLf
            Dim dict As Dictionary
            Set dict = value.value
            firstPass = True
            For Each v In dict.Keys
                If Not firstPass Then objStr = objStr & vbLf
                firstPass = False
                Dim key As String
                If TypeName(v) = "String" Then
                    key = CStr(v)
                Else ' assume pdfValue with valueType=PDF_Name
                    key = CStr(v.value)
                End If
                Set pv = dict.Item(key)
                objStr = objStr & key & " "
                objStr = objStr & BytesToString(pv.serialize(baseId))
            Next v
            If Right(objStr, 1) <> vbLf Then objStr = objStr & vbLf
            objStr = objStr & ">>" & vbLf
        Case PDF_ValueType.PDF_Stream       ' actual stream object with dictionary and data
            Dim stream As pdfStream
            Set stream = value.value
            IsBytes = True
            objBytes = stream.stream_meta.serialize(baseId)
            CopyBytes stream.stream_data.serialize(baseId), objBytes, 0, UBound(objBytes) + 1
        Case PDF_ValueType.PDF_StreamData   ' represents only stream ... endstream portion
            IsBytes = True
            objBytes = StringToBytes("stream" & vbLf)
            CopyBytes value.value, objBytes, 0, UBound(objBytes) + 1
            CopyBytes StringToBytes(vbLf & "endstream" & vbLf), objBytes, 0, UBound(objBytes) + 1
        
    ' to simplify processing, not one of 9 basic types either
        Case PDF_ValueType.PDF_Object       ' id generation obj << dictionary >> endobj
            IsBytes = True
            objBytes = StringToBytes(baseId + value.id & " " & value.generation & " obj" & vbLf)
            Dim pdfObj As pdfValue: Set pdfObj = value.value
            CopyBytes pdfObj.serialize(baseId), objBytes, 0, UBound(objBytes) + 1
            If objBytes(UBound(objBytes)) <> 10 Then CopyBytes StringToBytes(vbLf), objBytes, 0, UBound(objBytes) + 1
            CopyBytes StringToBytes("endobj" & vbLf), objBytes, 0, UBound(objBytes) + 1
        Case PDF_ValueType.PDF_Reference    ' indirect object
            ' Note: if indirect reference to /Parent and that obj is not in current set, will have wrong id, offset correct id by -baseId in Reference object prior to saving
            objStr = baseId + value.value & " " & value.generation & " R"
        Case PDF_ValueType.PDF_Comment
            objStr = value.value & vbLf
        Case PDF_ValueType.PDF_Trailer
            Dim pdfTrailer As pdfValue: Set pdfTrailer = value.value
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



' construction methods, these return new instances of pdfValue
' can be used by predeclared object, e.g. pdfValue.NewValueObj("/MyName", "/Name")

' returns a name as a pdfValue /Name obj
Function NewNameValue(ByVal name As String) As pdfValue
    Set NewNameValue = NewValue(name, "/Name")
End Function


' returns value as a pdfValue obj
' Note: if value is String then valueType can be used if want a PDF_Name or PDF_Trailer object instead of PDF_String
' a Dictionary returns a PDF_Dictionary and a Collection returns as PDF_Array
Function NewValue(ByRef value As Variant, Optional ByRef valueType As String = vbNullString, Optional ByVal id As Long = 0) As pdfValue
    Dim obj As pdfValue
    Set obj = New pdfValue
    
    ' initialize indirect reference object instance to be shared
    If id <> 0 Then
        obj.id = id
        Set obj.referenceObj = NewValue(obj, "/Reference")
    End If
    
    ' set obj's value and type information
    Select Case VarType(value)
        Case vbNull
            ' always return same instance for Null
            Static NullValue As pdfValue
            If NullValue Is Nothing Then
                obj.valueType = PDF_ValueType.PDF_Null
                Set NullValue = obj
            Else
                Set obj = NullValue
            End If
        Case vbLong, vbInteger
            obj.valueType = PDF_ValueType.PDF_Integer
            obj.value = CLng(value)
        Case vbSingle, vbDouble
            obj.valueType = PDF_ValueType.PDF_Real
            obj.value = CDbl(value)
        Case vbBoolean
            obj.valueType = PDF_ValueType.PDF_Boolean
            obj.value = CBool(value)
        Case vbString
            Select Case valueType
                Case "/Name"
                    obj.valueType = PDF_ValueType.PDF_Name
                    ' ensure /Name object begin with a /
                    If Left(value, 1) <> "/" Then value = "/" & value
                    obj.value = value
                Case Else
                    obj.valueType = PDF_ValueType.PDF_String
                    obj.value = CStr(value)
            End Select
        Case vbObject
            Select Case TypeName(value)
                Case "Dictionary"
                    obj.valueType = PDF_ValueType.PDF_Dictionary
                    Set obj.value = value
                Case "Collection"
                    obj.valueType = PDF_ValueType.PDF_Array
                    Set obj.value = value
                Case "pdfValue"
                    Select Case valueType
                        Case "/Trailer"
                            obj.valueType = PDF_ValueType.PDF_Trailer
                            Set obj.value = value
                        Case "/Reference"
                            obj.valueType = PDF_ValueType.PDF_Reference
                            obj.value = value.id
                            obj.generation = value.generation
                        Case Else ' default to "/Object"
                            obj.valueType = PDF_ValueType.PDF_Object
                            Set obj.value = value
                    End Select
                Case Else
                    Stop ' ???
            End Select
    End Select
        
    Set NewValue = obj
End Function

