VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pdfDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' simple API for reading, writing, and modifying Portable Document Format files
' Note: only smaller PDF documents that fully fit in memory supported,
'       as full document is read in at once, i.e. does not stream or load portions on demand
Option Explicit

' main pdf document structure
' header
' objects
' catalog
' trailer

Public rootCatalog As pdfValue  ' /Root object in cross reference table, document catalog, contains a pdf dictionary holding objects that define layout & contents of document
Public xrefTable As Dictionary  ' cross reference table consisting of xrefEntry key'd by object id
Public trailer As pdfValue      ' special pdf dictionary with references to key document objects such as /Root and /Info

' similar information, meta is newer but either or both may be present
Private m_Info As pdfValue
Public meta As pdfValue

' a PDF consists of a set of pages
Private m_Pages As pdfValue

' cache of loaded/shared objects
Public objectCache As Dictionary ' of pdfValue key'd by object id

' the full contents of raw PDF document
Private content() As Byte

' original filename and path
Public filename As String
Public filepath As String
' title
Private m_Title As String


Public Enum PDF_VERSIONS
    PDF_1_0 = 10
    PDF_1_1
    PDF_1_2
    PDF_1_3
    PDF_1_4
    PDF_1_5
    PDF_1_6
    PDF_1_7
    PDF_2_0 = 20
    PDF_Default = PDF_1_7
End Enum

' version in PDF
' generally one of: %PDF–1.0 %PDF–1.1 %PDF–1.2 %PDF–1.3 %PDF–1.4 %PDF–1.5 %PDF–1.6 %PDF–1.7 %PDF-2.0
' should be set to version string in document catalog if exists, otherwise version at start of file
Private m_version As PDF_VERSIONS
Private m_header As String

' update version, changes to version also update header
Public Property Let Version(ByVal pdfVersion As PDF_VERSIONS)
    m_version = pdfVersion
    m_header = "%PDF-" & Format(pdfVersion / 10#, "#.0") & vbNewLine
End Property
Public Property Get Version() As PDF_VERSIONS
    If m_version = 0 Then Me.Version = PDF_VERSIONS.PDF_Default   ' if not set, use default
    Version = m_version
End Property

' update header, changes to header also update version
Public Property Let Header(ByVal pdfHeader As String)
    m_header = pdfHeader
    Dim position As Long
    position = InStr(1, pdfHeader, "-", vbBinaryCompare)
    If position > 0 Then
        Dim s As String
        s = TrimWS(Mid(pdfHeader, position + 1))
        m_version = Round(CDbl(s), 1)
    Else
        m_version = PDF_Default
    End If
End Property
Public Property Get Header() As String
    If IsBlank(m_header) Then   ' get default header and version
        Dim defVersion As PDF_VERSIONS
        defVersion = Version    ' sets header when it initializes version
    End If
    Header = m_header
End Property


' get/set Info object; updates /Info in trailer
Public Property Set Info(ByRef Info As pdfValue)
    Set m_Info = Info
    
    ' update /Info in trailer
    If (Not trailer Is Nothing) Then
        Dim dict As Dictionary
        Set dict = trailer.asDictionary()
        ' handle /Info is Nothing or pdf null value
        Dim noInfoValue As Boolean  ' work around for VBA not having short circuit if
        If m_Info Is Nothing Then
            noInfoValue = True
        Else
            noInfoValue = (m_Info.valueType = PDF_ValueType.PDF_Null)
        End If
        If noInfoValue Then
            If dict.Exists("/Info") Then
                dict.Remove ("/Info")
            End If
        Else
            Set dict("/Info") = pdfValueObj(m_Info, "/Reference")
        End If
    End If
End Property
Public Property Get Info() As pdfValue
    Set Info = m_Info
End Property


' get/set Pages object
Public Property Set pages(ByRef pages As pdfValue)
    Set m_Pages = pages
End Property
Public Property Get pages() As pdfValue
    On Error GoTo errHandler
    If m_Pages Is Nothing Then
        ' find /Pages reference in /Root catalog and returns corresponding object from pdf or objectCache Dictionary
        Dim obj As pdfValue
        Dim dict As Dictionary
        Set dict = obj.Value
        If rootCatalog.asDictionary.Exists("/Pages") Then ' /Pages in the catalog dictionary << >>
            Set obj = rootCatalog.asDictionary.Item("/Pages")
            ' obj should now be a reference to our /Pages object
            If obj.valueType = PDF_ValueType.PDF_Reference Then
                ' retrieve object (not reference) matching given obj reference's (id, generation) pair
                Set m_Pages = getCachedObject(obj.Value, obj.generation)
            ElseIf obj.valueType = PDF_ValueType.PDF_Object Then
                ' ok, weird but whatever
                Set m_Pages = obj
            Else
                ' error! didn't find what we expected here
                Stop
            End If
        End If
    End If
    Set pages = m_Pages
    Exit Property
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Property


' returns matching object from PDF
' will first try to find in objectCache, then uses cross reference table to load from content,
' otherwise returns pdf null object
Public Function getCachedObject(ByVal id As Long, ByVal generation As Long) As pdfValue
    Dim obj As pdfValue
    Dim v As Variant
    For Each v In objectCache.Items
        Set obj = v
        If obj.id = id And obj.generation = generation Then
            Set getCachedObject = obj
            Exit For
        End If
    Next v
    ' fail to find so return null object
    Set getCachedObject = New pdfValue
End Function


' initializes pdfDocument by loading a PDF document, does not parse but does minimal validity checks
' True on success, false on any error
Public Function loadPdf(ByVal pdfFilename As String) As Boolean
    On Error GoTo errHandler
    ' save filename, but without any path
    If (InStr(1, pdfFilename, "\", vbBinaryCompare) > 0) Or _
       (InStr(1, pdfFilename, "/", vbBinaryCompare) > 0) Then
        Dim position1 As Long, position2 As Long
        position1 = InStr(1, pdfFilename, "\", vbBinaryCompare)
        position2 = InStr(1, pdfFilename, "/", vbBinaryCompare)
        If position1 >= position2 Then
            filename = Mid(pdfFilename, position1 + 1)
            filepath = Left(pdfFilename, position1)
        Else
            filename = Mid(pdfFilename, position2 + 1)
            filepath = Left(pdfFilename, position2)
        End If
    Else
        filename = pdfFilename
        filepath = vbNullString
    End If
    ' read in contents of file
    Dim fileLen As Long
    content = readFile(pdfFilename, fileLen)
    If fileLen < 1 Then
        MsgBox "Error reading in pdf!", vbOKOnly Or vbCritical, pdfFilename
        Exit Function
    End If
    If fileLen > 100& * 1024& * 1024& Then ' arbitrarily limit files to 100GB
        MsgBox "Warning pdf exceeds supported file size!", vbOKOnly Or vbCritical, pdfFilename
        Exit Function
    End If
    
    ' get start of xref
    'Dim xrefOffset As Long: xrefOffset = GetXrefOffset(content)
    
    ' get trailer with /Root information
    Set trailer = GetTrailer(content)
    
    ' load the xref table
    Set xrefTable = GetXrefTable(content, trailer)
    
    ' display info
    Set Info = GetInfoObject(content, trailer, xrefTable)
    Debug.Print BytesToString(Info.serialize())
    objectCache.Add Info.id, Info
    
    ' root obj of PDF
    Set rootCatalog = GetRootObject(content, trailer, xrefTable)
    Debug.Print BytesToString(rootCatalog.serialize())
    objectCache.Add rootCatalog.id, rootCatalog
    
    ' get metadata (replaces info but both may exist)
    If rootCatalog.asDictionary.Exists("/Metadata") Then
        Dim pdfRefObj As pdfValue
        Set pdfRefObj = rootCatalog.asDictionary.Item("/Metadata")
        Set meta = getObject(content, xrefTable, pdfRefObj.Value)
        Debug.Print BytesToString(meta.serialize())
        objectCache.Add meta.id, meta
    Else
        Debug.Print "No /Metadata object specified in root /Catalog"
    End If
    
    loadPdf = True ' success
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function


' attempts to parse PDF document filling objectCache
' will fail if loadPDF has not been called or content() array is not prefilled by other means
' True on success, false on any error
Public Function parsePdf() As Boolean
    GetObjectsInTree rootCatalog, content, xrefTable, objectCache
    parsePdf = True ' success
End Function


' writes this object as pdf file
' regenerates cross reference table and clears content()
Public Function savePdf() As Boolean
    If Not IsBlank(filename) Then savePdf = savePdfAs(filepath & filename)
End Function

' writes this object as pdf file as new file
' regenerates cross reference table and clears content()
Public Function savePdfAs(ByVal fileNameAndPath As String) As Boolean
    ' because it will no longer match our cross reference table, clear contents()
    On Error Resume Next
    Erase content()
    On Error GoTo errHandler
        
    Dim offset As Long
    Dim outputFileNum As Integer
    outputFileNum = SavePdfHeader(fileNameAndPath, offset, Header)
    
    Set xrefTable = NewXrefTable()  ' SavePdfObjects & SavePdfTrailer rebuild cross reference table with objects saved
    SavePdfObjects outputFileNum, objectCache, offset
    SavePdfTrailer outputFileNum, offset
    
    Close outputFileNum
    savePdfAs = True ' success
    Debug.Print "Saved " & fileNameAndPath
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Function



' SavePDF split into 3 parts

' writes out PDF header with version information %PDF-#.# to identify file is a PDF
' returns file handle to write rest of content to
' Warning: will delete file pdfFilename before attempting to open if it (file will no longer exist unless open succeeds)
Function SavePdfHeader(ByRef pdfFilename As String, ByRef offset As Long, Optional ByVal Header As String) As Integer
    ' delete if file exists, as otherwise may be extra junk at end of file, but ignore if doesn't exist or other error
    On Error Resume Next
    Kill pdfFilename
    On Error GoTo errHandler
    
    Dim outputFileNum As Integer
    outputFileNum = FreeFile
    Open pdfFilename For Binary Access Write Lock Write As #outputFileNum
    
    
    Const defHeader As String = "%PDF-1.7" & vbNewLine
    If IsBlank(Header) Then Header = defHeader
    offset = PutString(outputFileNum, Header)

    SavePdfHeader = outputFileNum
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Function

' save a single PDF object
' Warning: does not save stream object streams
Sub SavePdfObject(ByRef outputFileNum As Integer, ByRef obj As pdfValue, _
            ByRef offset As Long, Optional ByVal baseId As Long = 0, Optional ByVal prettyPrint As Boolean = True)
    On Error GoTo errHandler
    
    If False And (obj.valueType = PDF_ValueType.PDF_Stream) Then
        ' see if its a stream object stream (any other object stream  written out asis)
        If obj.asDictionary.Exists("/Type") Then
            Dim pdfNameValue As pdfValue
            Set pdfNameValue = obj.asDictionary.Item("/Type")
            If IsMatch(pdfNameValue.Value, "/ObjStm") Then Exit Sub
        End If
    End If
    
    If prettyPrint Then offset = offset + PutString(outputFileNum, vbLf)
    AddUpdateXref obj, offset, baseId
    offset = offset + PutBytes(outputFileNum, obj.serialize(baseId))
    
    Exit Sub
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Sub

' given current position in file, updates cross reference table and output serialized object
' for each object in pdfObjs dictionary.
' Note: assumes PDF Objects all have valid id values (not modified)
' Only writes objects provided, so if based on retrieving all objects from an
' arbitrary /Root catalog will only include referenced objects.
' Note: does not support compression (Deflate or other Filters), object streams
' are written out asis
' Warning: stream object streams are not written out, the contained objects are all
' written out as separate objects
Sub SavePdfObjects(ByRef outputFileNum As Integer, _
            ByRef pdfObjs As Dictionary, _
            ByRef offset As Long, _
            Optional ByVal baseId As Long = 0)
    On Error GoTo errHandler
    
    Dim v As Variant
    For Each v In pdfObjs.Items
        Dim obj As pdfValue: Set obj = v
        ' we skip saving /Info here as explicitly saved when saving trailer
        If obj.id <> Info.id Then
            SavePdfObject outputFileNum, obj, offset, baseId
        End If
    Next v
    
    Exit Sub
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Sub

' writes out info object, trailer, and cross reference table to given file handle beginning at offset
' Note: if info is pdf null value then skips writing /Info
' Note: this will also close the outputFileNum file handle
Sub SavePdfTrailer(ByRef outputFileNum As Integer, _
            ByRef offset As Long, _
            Optional ByVal prettyPrint As Boolean = True)
'            ByRef trailer As pdfValue,
'            ByRef xrefTable As Dictionary,
'            ByRef info As pdfValue,
'            ByRef root As pdfValue,
    On Error GoTo errHandler

    If Info.valueType <> PDF_ValueType.PDF_Null Then SavePdfObject outputFileNum, Info, offset, 0, prettyPrint
    If prettyPrint Then offset = offset + PutString(outputFileNum, vbLf)
    
    ' output xref catalog, for simple form, order should match id#s
    ' each entry should be exactly 20 bytes include 2 character whitespace so needs to end with \r\n or <space>\r or <space>\n
    ' Note: we may leave some id's unused, so we need to actually calculate our highest id and cycle through that in order
    Dim v As Variant
    Dim entry As xrefEntry
    Dim maxId As Long
    maxId = xrefTable.Count - 1 ' should be at least this high
    For Each v In xrefTable.Items
        Set entry = v
        If entry.id > maxId Then maxId = entry.id
        If entry.nextFreeId > maxId Then maxId = entry.nextFreeId
    Next v
    Dim xrefOffset As Long: xrefOffset = offset
    offset = offset + PutString(outputFileNum, "xref" & vbLf & "0 " & (maxId + 1) & vbNewLine)
    Dim ndx As Long
    For ndx = 0 To maxId ' Note we need to check actual id values for highest value and not just use xrefTable.count - 1
        If xrefTable.Exists(ndx) Then
            Set entry = xrefTable.Item(ndx)
            If entry.isFree Then
                PutString outputFileNum, Format(entry.nextFreeId, "0000000000") & " " & Format(entry.generation, "00000") & " f" & vbNewLine
            Else
                PutString outputFileNum, Format(entry.offset, "0000000000") & " " & Format(entry.generation, "00000") & " n" & vbNewLine
            End If
        Else
            'we could output new starting id# & count, but instead we use alternate of puting deleted item record
            PutString outputFileNum, "0000000000 00001 f" & vbNewLine
        End If
    Next ndx
    
    ' update our trailer with correct /Size of combined objects
    ' Note: we replace any existing /Size value in trailer (simpler logic than loading current value, checking if exists, creating if not, updating value)
    Dim obj As pdfValue
    Set obj = pdfValueObj(CLng(maxId + 1))
    Set trailer.asDictionary("/Size") = obj
    PutBytes outputFileNum, trailer.serialize(0)

    ' If prettyPrint then use vbNewLine, else use vbLf here
    If prettyPrint Then
        PutString outputFileNum, "startxref" & vbLf & xrefOffset & vbNewLine
        PutString outputFileNum, "%%EOF" & vbNewLine
    Else
        PutString outputFileNum, "startxref" & vbLf & xrefOffset & vbLf
        PutString outputFileNum, "%%EOF" & vbLf
    End If
    
    Close #outputFileNum
    Exit Sub
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Sub



' returns a default initialized trailer pdf object
' Note: as there are no objects yet, does not specify /Root reference
Function NewTrailer() As pdfValue
    Dim trailerDict As pdfValue
    Dim dict As Dictionary
    Set dict = New Dictionary
    dict.Add pdfValueObj("/Size", "/Name"), 1   ' the default empty slot
    Set trailerDict = pdfValueObj(dict)
    Set NewTrailer = pdfValueObj(trailerDict, "/Trailer")
End Function


' returns default cross reference table with required object id 0 free
Function NewXrefTable() As Dictionary
    Dim xrefTable As Dictionary
    Set xrefTable = New Dictionary
    Dim entry As xrefEntry
    Set entry = New xrefEntry
    entry.id = 0
    entry.generation = 65535
    entry.isFree = True
    entry.nextFreeId = 0
    entry.offset = 0
    xrefTable.Add entry.id, entry
    Set entry = Nothing
    Set NewXrefTable = xrefTable
    Set xrefTable = Nothing
End Function


' updates offset in existing cross reference table entry or adds new entry to cross reference table for obj
Sub AddUpdateXref(ByRef obj As pdfValue, ByVal offset As Long, Optional ByVal baseId As Long = 0)
    On Error GoTo errHandler
    Dim entry As xrefEntry
    Dim id As Long
    id = baseId + obj.id
    If xrefTable.Exists(id) Then
        If IsEmpty(xrefTable.Item(id)) Then
            Stop
            GoTo newEntry
        End If
        ' update existing entry
        Set entry = xrefTable.Item(id)
        entry.offset = offset
    Else
newEntry:
        Set entry = New xrefEntry
        entry.id = id
        entry.generation = obj.generation
        entry.isFree = False
        entry.nextFreeId = 0
        entry.offset = offset
        ' add to our catalog
        Set xrefTable(id) = entry
    End If
    Set entry = Nothing
    Exit Sub
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Sub



Private Sub Class_Initialize()
    Set objectCache = New Dictionary
    Set xrefTable = NewXrefTable()
    Set rootCatalog = New pdfValue
    Set trailer = NewTrailer()
    Set Info = New pdfValue
    Set meta = New pdfValue
End Sub

Private Sub Class_Terminate()
    Set objectCache = Nothing
    Set xrefTable = Nothing
    Set rootCatalog = Nothing
    Set trailer = Nothing
    Set m_Info = Nothing    ' don't set Info as we don't want to update trailer
    Set meta = Nothing
End Sub
