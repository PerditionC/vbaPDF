Attribute VB_Name = "pdfParseAndGetValues"
' parses and loads PDF values
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


' validates content() probably is a PDF document and returns version from PDF declaration
' returns True if found a valid PDF version, False on any error or invalid version declaration
' Note: if prepended extra data appears before header then it is ignored, offset is start of %PDF header
Function GetPdfHeader(ByRef content() As Byte, ByRef headerVersion As String, ByRef offset As Long) As Boolean
    On Error GoTo errHandler
    ' PDF documents should begin with something like %PDF-1.7<whitespace>
    Dim pdfHeaderVersion As String
    Dim fileSize As Long: fileSize = ByteArraySize(content)
    If fileSize < 10 Then Exit Function ' file is too small to be valid, not even %PDF-1.x\n
    ' scan from beginning
    offset = FindToken(content, "%PDF-")
    If offset < 0 Then Exit Function ' header not found!
    
    ' extract initial string in file, reusing GetWord but it should work fine for this purpose
    Dim headerOffset As Long: headerOffset = offset
    pdfHeaderVersion = TrimWS(GetWord(content, headerOffset))
    
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
Function GetDictionaryValue(ByRef value As pdfValue, ByVal name As String) As pdfValue
    On Error GoTo errHandler
    If value.value.Exists(name) Then
        Set GetDictionaryValue = value.value.Item(name)
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
    Set GetRoot = GetDictionaryValue(trailer.value, "/Root")
End Function


' returns /Info obj in PDF if exists, otherwise return PDF_Null
' expects to be given the pdfValue returned from GetTrailer() call
Function GetInfo(ByRef trailer As pdfValue) As pdfValue
    Set GetInfo = GetDictionaryValue(trailer.value, "/Info")
End Function


' returns how many xref entries in xref table (/Size value)
' expects to be given the pdfValue returned from GetTrailer() call
Function GetXrefSize(ByRef trailer As pdfValue) As Long
    On Error GoTo errHandler
    Dim Size As pdfValue
    Set Size = GetDictionaryValue(trailer.value, "/Size")
    If Size.valueType = PDF_ValueType.PDF_Integer Or Size.valueType = PDF_ValueType.PDF_Real Then
        GetXrefSize = CLng(Size.value)
    End If
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Function
