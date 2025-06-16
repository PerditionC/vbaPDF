VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pdfDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

' PDF 2.0 allow arbitrary data prior to PDF header, with all offsets beginning at start of header (the %PDF)
' when reading, if there is data we store it here and rewrite it on save, but it is otherwise unused
Public prependedData As Variant

' similar information, meta is newer but either or both may be present
Private m_Info As pdfValue
Private m_Meta As pdfValue

' a PDF consists of a set of pages
' use pages.referenceObj where any /Parent or other indirect reference object is needed
Private m_Pages As pdfValue

' cache of loaded/shared objects
Public objectCache As Dictionary ' of pdfValue key'd by object id

' cache only used when loading embedded objects from stream object streams, caches the stream object
Private sosCache As Dictionary
' used to maintain singleton instances of all direct & indirect reference objects (so can correctly track object id if it changes)
Private refCache As Dictionary

' the full contents of raw PDF document
Private content() As Byte

' original filename and path
Public filename As String
Public filepath As String
' title
Private m_Title As String


' first free object id that can be used, access via nextObjId Property
Private m_id As Long


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
' generally one of: %PDF 1.0 %PDF 1.1 %PDF 1.2 %PDF 1.3 %PDF 1.4 %PDF 1.5 %PDF 1.6 %PDF 1.7 %PDF-2.0
' should be set to version string in document catalog if exists, otherwise version at start of file
Private m_version As PDF_VERSIONS
Private m_header As String

' update version, changes to version also update header
Public Property Let version(ByVal pdfVersion As PDF_VERSIONS)
    m_version = pdfVersion
    m_header = "%PDF-" & Format(pdfVersion / 10#, "#.0") & vbNewLine
End Property
Public Property Get version() As PDF_VERSIONS
    If m_version = 0 Then Me.version = PDF_VERSIONS.PDF_Default   ' if not set, use default
    version = m_version
End Property

' update header, changes to header also update version
Public Property Let Header(ByVal pdfHeader As String)
    m_header = pdfHeader
    Dim position As Long
    position = InStr(1, pdfHeader, "-", vbBinaryCompare)
    If position > 0 Then
        Dim s As String
        s = TrimWS(Mid(pdfHeader, position + 1))
        m_version = Round(CDbl(s), 1) * 10
    Else
        m_version = PDF_Default
    End If
End Property
Public Property Get Header() As String
    If IsBlank(m_header) Then   ' get default header and version
        Dim defVersion As PDF_VERSIONS
        defVersion = version    ' sets header when it initializes version
    End If
    Header = m_header
End Property


' returns the next available object id
Public Property Let nextObjId(ByVal nextId As Long)
    m_id = nextId
End Property
Public Property Get nextObjId() As Long
    m_id = m_id + 1
    nextObjId = m_id
End Property


' updates object reference in trailer, does nothing if trailer is not valid
Private Sub updateTrailerReference(ByRef key As String, ByRef valueObj As pdfValue)
    If (Not trailer Is Nothing) Then
        Dim dict As Dictionary
        Set dict = trailer.asDictionary()
        ' handle /Info is Nothing or pdf null value
        Dim noValue As Boolean  ' work around for VBA not having short circuit if
        If valueObj Is Nothing Then
            noValue = True
        Else
            noValue = (valueObj.valueType = PDF_ValueType.PDF_Null)
        End If
        If noValue Then
            If dict.Exists(key) Then
                dict.Remove key
            End If
        Else
            Set dict(key) = valueObj.referenceObj
        End If
    End If
End Sub

' get/set Info object; updates /Info in trailer
' note: /Metadata is newer method and replaces /Info, however, both may exist
' /Metadata stored in /Catalog dictionary in XML while /Info is dictionary object stored in /Trailer
' Will first attempt to load from objectCache (parsePDF previously called)
' otherwise attempts to parse directly from raw content() -- does not add to objectCache
' Expects trailer object to already be set/loaded (e.g. by loadPDF)
Public Property Set Info(ByRef Info As pdfValue)
    Set m_Info = Info
    
    ' update /Info in trailer
    updateTrailerReference "/Info", m_Info
    
    ' update our Title on next access
    m_Title = vbNullString  ' we replaced /Info, so that changes Title too
End Property
Public Property Get Info() As pdfValue
    If m_Info Is Nothing Then Set m_Info = New pdfValue
    If m_Info.valueType = PDF_ValueType.PDF_Null Then
        ' 1st try loading cached object, if that fails, try loading from raw content()
        ' note: we Set m_Info instead of Me.Info as we are retrieving /Info id from trailer's reference object already
    
        ' get reference to /Info object (object itself should not be stored directly in trailer)
        ' if no /Info key or it refers to null, then returns pdf null value
        ' (if /Info stored directly as object, not valid, but we will return that anyway)
        Dim infoRefObj As pdfValue: Set infoRefObj = GetInfo(trailer)
        If infoRefObj.valueType = PDF_ValueType.PDF_Reference Then
            Set m_Info = getObject(CLng(infoRefObj.value), CLng(infoRefObj.generation), cacheObject:=False)
        'ElseIf info.valueType = PDF_ValueType.PDF_Object Then
        'ElseIf info.valueType = PDF_ValueType.PDF_Null Then
        'Else error, unexpected type
        End If
        
        Debug.Print "/Info=" & BytesToString(m_Info.serialize())
    End If

    Set Info = m_Info
End Property


' returns either /Title in /Info or if not found (no /Info or no /Title, filename)
' Note: we do not update /Meta as Meta is XML which we don't currently want to deal with
Public Property Let Title(pdfTitle As String)
    Stop ' TODO update/add /Info
End Property
Public Property Get Title() As String
    If IsBlank(m_Title) Then
        If Info.hasKey("/Title") Then
            m_Title = Info.asDictionary().Item("/Title").value
        Else
            m_Title = filename
        End If
    End If
    
    Title = m_Title
End Property


' get/set Meta object; updates /Metadata in trailer
' note: /Metadata is newer method and replaces /Info, however, both may exist
' /Metadata stored in /Catalog dictionary in XML while /Info is dictionary object stored in /Trailer
' Will first attempt to load from objectCache (parsePDF previously called)
' otherwise attempts to parse directly from raw content() -- does not add to objectCache
' Expects rootCatalog or trailer object to already be set/loaded (e.g. by loadPDF)
Public Property Set Meta(ByRef Meta As pdfValue)
    Set m_Meta = Meta

    ' update /Info in trailer
    If (Not rootCatalog Is Nothing) Then
        Dim dict As Dictionary
        Set dict = rootCatalog.asDictionary()
        If Not dict Is Nothing Then
            ' handle /Metadata is Nothing or pdf null value
            Dim noMetaValue As Boolean  ' work around for VBA not having short circuit if
            If m_Meta Is Nothing Then
                noMetaValue = True
            Else
                noMetaValue = (m_Meta.valueType = PDF_ValueType.PDF_Null)
            End If
            If noMetaValue Then
                If dict.Exists("/Metadata") Then
                    dict.Remove ("/Metadata")
                End If
            Else
                Set dict("/Metadata") = m_Meta.referenceObj
            End If
        End If
    End If
End Property
Public Property Get Meta() As pdfValue
    If m_Meta Is Nothing Then Set m_Meta = New pdfValue
    If m_Meta.valueType = PDF_ValueType.PDF_Null Then
        ' /Metadata is part of rootCatalog dictionary
        Dim catalogDict As Dictionary
        ' has rootCatalog been established yet? if not attempt to get it
        If rootCatalog.valueType = PDF_ValueType.PDF_Null Then
            ' root obj of PDF
            Dim root As pdfValue ' we are not loading rootCatalog here, user must explicity set themselves or call loadPDF
            Set root = GetRootObject(content, trailer, xrefTable)
            'Debug.Print BytesToString(root.serialize())
            Set catalogDict = root.asDictionary()
            Set root = Nothing
        Else
            Set catalogDict = rootCatalog.asDictionary()
        End If
        ' objtain reference to /Metadata object and get it
        If Not catalogDict Is Nothing Then
            If catalogDict.Exists("/Metadata") Then
                Dim pdfRefObj As pdfValue
                Set pdfRefObj = catalogDict.Item("/Metadata")
                ' 1st try loading cached object, if that fails, try loading from raw content()
                Set m_Meta = getObject(CLng(pdfRefObj.value), pdfRefObj.generation, cacheObject:=False)
                Debug.Print BytesToString(m_Meta.serialize())
            Else
                Debug.Print "No /Metadata object specified in root /Catalog"
            End If
        Else
            Debug.Print "No root /Catalog found, no /Metadata object"
        End If
    End If

    Set Meta = m_Meta
End Property


' get/set Pages object
Public Property Set pages(ByRef pages As pdfValue)
    Set m_Pages = pages
End Property
Public Property Get pages() As pdfValue
    On Error GoTo errHandler
    If m_Pages Is Nothing Then Set m_Pages = New pdfValue
    If m_Pages.valueType = PDF_ValueType.PDF_Null Then
        ' find /Pages reference in /Root catalog and returns corresponding object from pdf or objectCache Dictionary
        If rootCatalog.hasKey("/Pages") Then ' /Pages in the catalog dictionary << >>
            Dim obj As pdfValue
            Set obj = rootCatalog.asDictionary.Item("/Pages")
            ' obj should now be a reference to our /Pages object
            If obj.valueType = PDF_ValueType.PDF_Reference Then
                ' retrieve object (not reference) matching given obj reference's (id, generation) pair
                Set m_Pages = getObject(CLng(obj.value), obj.generation)
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


' returns /Count from /Pages (total /Page objects in pdf)
' readonly - add/remove /Page objects and update /Pages tree to alter value
Public Property Get pageCount() As Long
    If pages.hasKey("/Count") Then
        pageCount = CLng(pages.asDictionary().Item("/Count").value)
    End If
End Property


' helper methods when creating a new pdf document or merging multiple ones

' adds & initializes the /Trailer/Info dictionary
' Note: if /Info already exists in /Trailer, this will replace it with a new one, assumes trailer is valid
Public Sub AddInfo(Optional defaults As Variant)
    Set Info = pdfValue.NewValue(pdfValue.NewValue(New Dictionary), id:=nextObjId) ' also updates /Trailer
    
    ' add in some additional default values
    With Info.asDictionary()
        Set .Item("/Producer") = pdfValue.NewValue("vbaPDF")
    End With
End Sub

' adds & initializes the top level /Root document catalog
' Note: if /Root already exists in /Trailer, this will replace it with a new one, assumes trailer is valid
Public Sub NewDocumentCatalog()
    Set rootCatalog = pdfValue.NewValue(pdfValue.NewValue(New Dictionary), id:=nextObjId)
    ' add in some additional default values
    With rootCatalog.asDictionary()
        Set .Item("/Type") = pdfValue.NewNameValue("/Catalog")
        Set .Item("/Version") = pdfValue.NewNameValue("/" & Format(version / 10#, "#.0"))
        Set .Item("/PageLayout") = pdfValue.NewNameValue("/OneColumn")  ' if not specified default is "/SinglePage"
        Set .Item("/PageMode") = pdfValue.NewNameValue("/UseNone")      '  " is "/UseNone"
        Set .Item("/Lang") = pdfValue.NewValue("en")                    '  " is unknown language
    End With
    
    updateTrailerReference "/Root", rootCatalog
End Sub

' adds a /Page or /Pages to document
' if pages is Nothing then will also set top level /Root /Pages to newPages
Public Sub AddPages(Optional ByRef thePages As pdfValue = Nothing)
    ' ensure top level /Pages is initialized
    If pages.valueType = PDF_ValueType.PDF_Null Then
        Set pages = NewPages(Nothing)
    End If
    
    ' add new /Pages or /Page to our top level /Pages
    If Not thePages Is Nothing Then
        ' first get Kids
        Dim kids As Collection
        ' add our Page(s) as a new child (as an indirect object reference)
        Set kids = pages.asDictionary("/Kids").value
        kids.Add thePages.referenceObj
        Set thePages.asDictionary("/Parent") = pages.referenceObj
        ' update page count
        Set pages.asDictionary("/Count") = pdfValue.NewValue(pageCount + CLng(thePages.asDictionary("/Count").value))
    End If
End Sub
' creates a default initialized /Page object
Public Function NewPage(Optional defaults As Variant) As pdfValue

End Function
' creates a default initialized /Pages object
Public Function NewPages(ByRef Parent As pdfValue, Optional defaults As Variant) As pdfValue
    Set NewPages = pdfValue.NewValue(pdfValue.NewValue(New Dictionary), id:=nextObjId)
    ' add in some additional default values
    With NewPages.asDictionary()
        Set .Item("/Type") = pdfValue.NewNameValue("/Pages")
        ' /Parent should not exist in /Root/Pages, but need to update /Root to point to our new toplevel /Pages
        If Parent Is Nothing Then
            Set rootCatalog.asDictionary("/Pages") = NewPages.referenceObj
        Else
            If Parent.valueType = PDF_ValueType.PDF_Object Then
                Set .Item("/Parent") = Parent.referenceObj
            ElseIf Parent.valueType = PDF_ValueType.PDF_Reference Then
                Set .Item("/Parent") = Parent ' already a reference
            Else
                Stop ' unsupported value!
            End If
        End If
        Set .Item("/Kids") = pdfValue.NewValue(New Collection)
        Set .Item("/Count") = pdfValue.NewValue(0)
    End With
    
    ' ensure /Pages object is saved and easily retrievable
    objectCache.Add NewPages.id, NewPages
End Function

' TODO
Public Sub AddOutlines()

End Sub


' updates id of all objects under root pdfValue beginning at baseId
' Note: use trailer as root to renumber whole document
' Warning: assumes each obj id refers to same obj and objReference instances
Public Function renumberIds(ByRef root As pdfValue, ByVal baseId As Long, Optional ByRef visited As Dictionary = Nothing) As Long
    On Error GoTo errHandler
    If visited Is Nothing Then Set visited = New Dictionary
    Dim v As Variant
    Dim obj As pdfValue
    ' sanity check
    If root Is Nothing Then Exit Function
    ' don't renumber same object more than once
    If root.id <> 0 Then
        If visited.Exists(root.id) Then
            renumberIds = baseId
            Exit Function
        End If
        ' and actually do the renumbering
        root.id = baseId
        baseId = baseId + 1
        visited.Add root.id, root.id
    End If
    ' update any children
    Select Case root.valueType
        Case PDF_ValueType.PDF_Array
            For Each v In root.value
                Set obj = v
                baseId = renumberIds(obj, baseId, visited)
            Next v
        Case PDF_ValueType.PDF_Dictionary
            For Each v In root.value.Items
                Set obj = v
                baseId = renumberIds(obj, baseId, visited)
            Next v
        Case PDF_ValueType.PDF_Reference
            Set obj = getObject(root.value, 0)
            baseId = renumberIds(obj, baseId, visited)
        Case PDF_ValueType.PDF_Trailer, PDF_ValueType.PDF_Object
            baseId = renumberIds(root.value, baseId, visited)
        Case Else
            ' no nested objects that need id's updated
    End Select
    ' last id used, or same as provided if nothing updated
    renumberIds = baseId
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function



' returns matching object from PDF
' will first try to find in objectCache, then uses cross reference table to load from content(),
' otherwise returns pdf null object
' Note: unless cacheObject is set to False, if successfully extracted object from content() then will add to object cache
Public Function getObject(ByVal id As Long, ByVal generation As Long, Optional ByVal cacheObject As Boolean = True) As pdfValue
    Dim obj As pdfValue
    ' first try loading from our cache
    Set obj = getCachedObject(id, generation)
    ' if it wasn't there then try loading from content()
    If obj.valueType = PDF_ValueType.PDF_Null Then
        Set obj = loadObject(content, xrefTable, id)
        If obj Is Nothing Then Set obj = New pdfValue
        
        ' if successfully loaded, update our object cache
        If cacheObject And (obj.valueType <> PDF_ValueType.PDF_Null) Then
            If objectCache.Exists(id) Then
                ' we need to determine if same generation, presumably not or we would have retrieved it above
                Dim dict As Dictionary
                Dim curObj As pdfValue
                Dim v As Variant
                Set v = objectCache.Item(id)
                If TypeName(v) = "pdfValue" Then
                    Set curObj = objectCache.Item(id)
                    If curObj.generation <> generation Then
                        Set dict = New Dictionary
                        dict.Add curObj.generation, curObj
                        dict.Add generation, obj
                        Set objectCache(id) = dict
                        Set dict = Nothing
                    Else
                        Debug.Print "Internal Error! - id & generation the same but not?"
                        Stop
                    End If
                    Set curObj = Nothing
                ElseIf TypeName(v) = "Dictionary" Then
                    ' assume may contain multiple generations, key'd by generation
                    Set dict = v
                    If dict.Exists(generation) Then
                        Set dict(generation) = obj
                    Else
                        dict.Add generation, obj
                    End If
                    Set dict = Nothing
                Else
                    Debug.Print "Internal Error!"
                    Stop
                End If
            Else
                objectCache.Add id, obj
            End If
        End If
    End If
    
    Set getObject = obj
End Function


' returns matching object from PDF
' will first try to find in objectCache, if not found in cache then returns pdf null object
' Does not attempt to load from raw pdf in content() via cross reference table
Public Function getCachedObject(ByVal id As Long, ByVal generation As Long) As pdfValue
    Dim v As Variant
    ' return object, but verify matches expected generation
    If objectCache.Exists(id) Then
        Set v = objectCache.Item(id)
        If TypeName(v) = "pdfValue" Then
            Set getCachedObject = v
            If getCachedObject.generation = generation Then Exit Function
        ElseIf TypeName(v) = "Dictionary" Then
            ' assume may contain multiple generations, key'd by generation
            If v.Exists(generation) Then
                Set getCachedObject = v.Item(generation)
                Exit Function
            End If
        Else
            Debug.Print "Internal Error!"
            Stop
        End If
    End If
    ' fail to find so return null object
    Set getCachedObject = New pdfValue
End Function


' initializes pdfDocument by loading a PDF document, does not parse but does minimal validity checks
' True on success, false on any error
' Note: baseId is used to renumber objs as loaded
Public Function loadPdf(ByVal pdfFilename As String) As Boolean
    On Error GoTo errHandler
    
    ' save filename, but without any path
    If (InStr(1, pdfFilename, "\", vbBinaryCompare) > 0) Or _
       (InStr(1, pdfFilename, "/", vbBinaryCompare) > 0) Then
        Dim position1 As Long, position2 As Long
        position1 = InStrRev(pdfFilename, "\", -1, vbBinaryCompare)
        position2 = InStrRev(pdfFilename, "/", -1, vbBinaryCompare)
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
    
    ' validate probably a PDF file and get header version (more recent version# may be specified in root Catalog)
    ' PDF 2.0 allows
    Dim headerVersion As String ' Note: we have to pass a String object as Me.Header won't save changes
    Dim offset As Long
    If Not GetPdfHeader(content, headerVersion, offset) Then
        Erase content: content = vbNullString
        Exit Function ' return False
    End If
    ' if funk before header move it and shrink our data
    If offset > 0 Then
        Dim tmpBuffer() As Byte
        ReDim tmpBuffer(0 To (offset - 1)) As Byte
        CopyBytes content, tmpBuffer, 0, 0, offset
        prependedData = tmpBuffer
        ReDim tmpBuffer(0 To (UBound(content) - offset))
        CopyBytes content, tmpBuffer, offset, 0
        content = tmpBuffer
    End If
    Me.Header = headerVersion   ' Note: this also updates our Version value
    
    ' get trailer with /Root information
    Set trailer = GetTrailer(content)
    
    ' load the xref table
    Set xrefTable = GetXrefTable(content, trailer)
    
    ' root obj of PDF
    Set rootCatalog = GetRootObject(content, trailer, xrefTable)
    Debug.Print BytesToString(rootCatalog.serialize())
    objectCache.Add rootCatalog.id, rootCatalog
    
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
    On Error GoTo errHandler
    GetObjectsInTree rootCatalog, content, xrefTable, objectCache
    ' we need to get current highest object id used
    Dim v As Variant
    For Each v In objectCache.Items
        Dim id As Long: id = CLng(v.id)
        If id >= m_id Then nextObjId = id
    Next v
    
    ' /Info is in trailer, so won't be loaded as part of rootCatalog,
    ' Note first call to Info actually retreives, so no need to: Set Info = GetInfoObject(content, trailer, xrefTable)
    If Info.id <> 0 Then
        objectCache.Add Info.id, Info
    End If
    If Info.id >= m_id Then nextObjId = Info.id
    
    ' /Metadata is in root catalog so should be in object cache, but not Me.meta yet, make any call to initalize it
    If Meta.valueType <> PDF_ValueType.PDF_Null Then
        'Debug.Print BytesToString(Meta.serialize())
    End If
    
    parsePdf = True ' success
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
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
    If Err.Number = 70 Then ' Permission denied
        If (6 = MsgBox("File in use or permission denied, close file if open and retry?", vbYesNo Or vbQuestion, pdfFilename)) Then Resume ' 6=Yes
    Else
        Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
        Stop
    End If
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
            If IsMatch(pdfNameValue.value, "/ObjStm") Then Exit Sub
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
    Set trailer.asDictionary("/Size") = pdfValue.NewValue(CLng(maxId + 1))
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
    dict.Add "/Size", pdfValue.NewValue(1)    ' the default empty slot
    Set trailerDict = pdfValue.NewValue(dict)
    Set NewTrailer = pdfValue.NewValue(trailerDict, "/Trailer")
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
    On Error GoTo errHandler
    Set objectCache = New Dictionary
    Set sosCache = New Dictionary
    Set refCache = New Dictionary
    Set xrefTable = NewXrefTable()
    Set trailer = NewTrailer()
    NewDocumentCatalog
    Set Info = New pdfValue ' default is null
    Set Meta = New pdfValue ' default is null
    Exit Sub
errHandler:
    Debug.Print Err.Description
    Stop
    Resume
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set objectCache = Nothing
    Set sosCache = Nothing
    Set refCache = Nothing
    Set xrefTable = Nothing
    Set rootCatalog = Nothing
    Set trailer = Nothing
    Set m_Info = Nothing    ' don't set Info as we don't want to update trailer
    Set Meta = Nothing
    Erase content
    Exit Sub
errHandler:
    Debug.Print Err.Description
    Stop
    Resume Next
End Sub


' Warning: Only below methods are intended to be used by predeclared pdfDocument object
' Use of any other methods is unsupported (as they expect to work only on current instance)

' return a New instance of pdfDocument
' Optional, if filename is provided, will load and parse the PDF document
' Warning: if error loading or parsing, will return Nothing instead of partially loaded document
Public Function pdfDocument(Optional ByVal filename As String = vbNullString) As pdfDocument
    Set pdfDocument = New pdfDocument
    
    If Not IsBlank(filename) Then
        If pdfDocument.loadPdf(filename) Then
            If pdfDocument.parsePdf() Then
                Exit Function
            End If
        End If
        ' error loading, so return Nothing
        Set pdfDocument = Nothing
    End If
End Function



''''''''''''''''''

' given raw PDF file contents as Byte array and offset in the array to peak, returns value type at offset
Private Function GetValueType(ByRef bytes() As Byte, ByVal offset As Long) As PDF_ValueType
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
    Resume
End Function


' a Name may have value encoded in #00 form, ie # followed by 2 hex digits
Private Function ProcessName(ByRef name As String) As String
    Dim ndx As Long
    ndx = InStr(1, name, "#", vbBinaryCompare)
    If ndx > 0 Then
        Dim s As String, n As Long
        ProcessName = Left(name, ndx - 1)
        s = Mid(name, ndx + 1, 2) ' get hex digits
        n = CLng("&H" & s) ' convert from hex to Long
        s = ProcessName & Chr(n) & Mid(name, ndx + 3) ' combine with hex digit as character
        ProcessName = ProcessName(s) ' for now just recursively call to handle multiple encoded values, really should be a loop!
    Else
        ProcessName = name
    End If
End Function


' returns a value loaded for a PDF
' updates offset to next non-whitespace byte after this value is loaded
' bytes() may be content() or uncompressed stream object stream buffer
' offset is within bytes() provided
' Note: meta is only used for stream object
Function GetValue(ByRef bytes() As Byte, ByRef offset As Long, Optional ByRef Meta As Dictionary = Nothing) As pdfValue
    On Error GoTo errHandler
    DoEvents
    Dim result As pdfValue: Set result = New pdfValue
    result.id = 0
    result.generation = 0
    result.value = Empty
    result.valueType = GetValueType(bytes, offset)
    If offset > UBound(bytes) Then Exit Function        ' return null type if end of daa
    
    offset = SkipWhiteSpace(bytes, offset)
    If offset > UBound(bytes) Then Exit Function    ' return null type if end of data
    Dim token As String: token = Chr(bytes(offset))
    
    Dim tmpStr As String
    Dim name As pdfValue, value As pdfValue
    Select Case result.valueType
        Case PDF_ValueType.PDF_Null
            ' result.value = Empty
            If Not IsMatch(GetWord(bytes, offset), "null") Then
                Stop ' error, expecting "null"
            End If
        Case PDF_ValueType.PDF_Name
            result.value = ProcessName(GetWord(bytes, offset))
        Case PDF_ValueType.PDF_Boolean
            tmpStr = GetWord(bytes, offset)
            result.value = CBool(tmpStr)
        Case PDF_ValueType.PDF_Integer
            tmpStr = GetWord(bytes, offset)
            result.value = CLng(tmpStr)
        Case PDF_ValueType.PDF_Real
            tmpStr = GetWord(bytes, offset)
            result.value = CDbl(tmpStr)
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
                ' only need to escape unbalanced ), so need to keep track of balanced ()
                ' Note: could see ((\)) or (\(\))
                Dim openParenthesis As Integer
                Do While (openParenthesis > 0) Or (bytes(offset) <> 41) ' Asc(")") - if not open (, then end of string found
                    ' otherwise if ( or ) then update our tracking, but otherwise treat as normal character in string
                    If (bytes(offset) = 40) Then ' (
                        openParenthesis = openParenthesis + 1
                    ElseIf (bytes(offset) = 41) Then ' )
                        openParenthesis = openParenthesis - 1
                    End If
                    
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
                                openParenthesis = openParenthesis - 1
                            Case 40  ' Asc("(")
                                strBuffer(strLen) = 40   ' "("
                                strLen = strLen + 1
                                openParenthesis = openParenthesis + 1
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
            result.value = tmpStr
        Case PDF_ValueType.PDF_Array
            offset = offset + 1 ' skip the opening [
            Dim col As Collection
            Set col = New Collection
            Set result.value = col
            Set value = GetValue(bytes, offset)
            Do While value.valueType <> PDF_ValueType.PDF_EndOfArray
                col.Add value
                Set value = GetValue(bytes, offset)
            Loop
            Set col = Nothing
        Case PDF_ValueType.PDF_Dictionary
            offset = offset + 2 ' skip past <<
            Dim dict As Dictionary
            Set dict = New Dictionary
            Set result.value = dict
            Set name = GetValue(bytes, offset)
            Do While name.valueType = PDF_ValueType.PDF_Name
                Set value = GetValue(bytes, offset)
                dict.Add CStr(name.value), value
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
                If IsObject(Meta.Item("/Length").value) Then
                    ' Reference to actual /Length value, allowed but not common
                    Stop
                    ' TODO pass we need Dictionary of loaded objects or xrefTable to load value
                Else
                    ' /Length directly in file as number, usual
                    byteLen = CLng(Meta.Item("/Length").value)
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
            result.value = dataBytes
            If Not IsMatch(Left(tmpStr, 9), "endstream") Then Stop ' error unexpected token found!
            
        Case PDF_ValueType.PDF_Reference, PDF_ValueType.PDF_Object
            Dim words(0 To 2) As String
            words(0) = GetWord(bytes, offset)
If IsMatch(words(0), "3") Then Stop
            offset = SkipWhiteSpace(bytes, offset)
            words(1) = GetWord(bytes, offset)
            offset = SkipWhiteSpace(bytes, offset)
            words(2) = GetWord(bytes, offset)
            offset = SkipWhiteSpace(bytes, offset)
            If IsMatch(words(2), "obj") Then
                result.id = CLng(words(0))
                Set result.value = GetValue(bytes, offset)
                Dim endObjOrStream As pdfValue
                If result.value.valueType = PDF_ValueType.PDF_Dictionary Then
                    Set Meta = result.value.value
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
                    streamObj.Init result.value, endObjOrStream
                    Set stream.value = streamObj
                    Set streamObj = Nothing
                    Set result.value = stream
                    Set stream = Nothing
                    Set endObjOrStream = GetValue(bytes, offset)
                End If
                If endObjOrStream.valueType <> PDF_ValueType.PDF_EndOfObject Then
                    Stop ' error, expected "endobj"
                End If
            Else ' ISMatch(word(2), "R") Then
                result.value = CLng(words(0))
            End If
            result.generation = CLng(words(1))
        
            ' ensure same indirect reference object instance returned for same object always returned
            Dim singletonKey As String
            Select Case result.valueType
                Case PDF_ValueType.PDF_Object
                    If result.id <> 0 Then
                        ' ensure only 1 VBA instance of pdf obj
                        singletonKey = result.id & "_" & result.generation ' object key
                        If refCache.Exists(singletonKey) Then
                            Set result = refCache(singletonKey) ' use previously created instance
                        Else
                            Set refCache(singletonKey) = result ' store instance for later use
                        End If
                        ' ensure only 1 VBA instance of pdf indirect obj reference
                        With result
                            singletonKey = .id & " " & .generation & " R"  ' id gen R
                        End With
                        If refCache.Exists(singletonKey) Then
                            Set result.referenceObj = refCache(singletonKey)
                        Else
                            Set refCache(singletonKey) = result.referenceObj
                        End If
                    End If
                Case PDF_ValueType.PDF_Reference
                    If result.value <> 0 Then
                        ' Note: we don't preload obj as /Parent references can cause infinite loops & caller may not need obj yet
                        ' ensure only 1 VBA instance of pdf indirect obj reference
                        With result
                            singletonKey = .value & " " & .generation & " R"  ' id gen R
                        End With
                        If refCache.Exists(singletonKey) Then
                            Set result = refCache(singletonKey)
                        Else
                            Set refCache(singletonKey) = result
                        End If
                    End If
                Case Else
                    ' nothing to do
            End Select
    
        
        Case PDF_ValueType.PDF_Comment
            result.value = GetLine(bytes, offset)
        Case PDF_ValueType.PDF_Trailer
            If Not IsMatch(GetWord(bytes, offset), "trailer") Then
                Stop ' error, expected "trailer"
            End If
            Set result.value = GetValue(bytes, offset)
        Case PDF_ValueType.PDF_EndOfArray, PDF_ValueType.PDF_EndOfDictionary, PDF_ValueType.PDF_EndOfObject, PDF_ValueType.PDF_EndOfStream
            tmpStr = GetWord(bytes, offset) ' skip past end marker
        Case Else
            Stop ' error, unexpected type!
    End Select
    
    ' skip past any trailing whitespace
    offset = SkipWhiteSpace(bytes, offset)
    ' set our value here as result may now be a different instance (singleton value)
    Set GetValue = result
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
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


' extracts/parses pdf object from raw pdf content()
' due to potential slowness uncompressing in VBA, stream object streams should be cached
' the sosCache is only used for stream object streams and only if provided
' the refCache is to ensure only 1 indirect reference object instance is created for each object
' this will use provided VBA reference object instance instead of creating a new [different] one
Function loadObject(ByRef content() As Byte, ByRef xrefTable As Dictionary, ByVal Index As Long) As pdfValue
    On Error GoTo errHandler
    Dim obj As pdfValue
    If xrefTable.Exists(Index) Then
        Dim entry As xrefEntry
        Set entry = xrefTable.Item(Index)
        If entry.isFree Or ((Not entry.isEmbeded) And (entry.offset <= 0)) Then GoTo NullValue
        Dim offset As Long: offset = entry.offset
        If offset > UBound(content) Then GoTo NullValue
        If entry.isEmbeded Then
            Dim cntrObjEntry As xrefEntry
            Set cntrObjEntry = xrefTable.Item(entry.embedObjId)
            Dim cntrObj As pdfValue
            ' try loading containing object (stream object stream) from cache before potentially uncompressing
            If Not sosCache Is Nothing Then
                If sosCache.Exists(entry.embedObjId) Then Set cntrObj = sosCache(entry.embedObjId)
            End If
            If cntrObj Is Nothing Then  ' not in cache or no cache provided
                Set cntrObj = loadObject(content, xrefTable, entry.embedObjId)
                If Not sosCache Is Nothing Then Set sosCache(entry.embedObjId) = cntrObj    ' add/update cache
            End If
            
            ' extract our embedded object
            If cntrObj.value.valueType <> PDF_ValueType.PDF_Stream Then
                Debug.Print "Error! expecting stream object stream!"
                Stop
                GoTo NullValue
            End If
            Dim streamObjectStream As pdfStream
            Set streamObjectStream = cntrObj.value.value
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
                    firstOffset = CLng(dict.Item("/First").value)
                Else
                    firstOffset = -1
                End If
                Dim n As Long, id As Long, objOffset As Long
                Dim v As pdfValue
                If dict.Exists("/N") Then
                    n = CLng(dict.Item("/N").value)
                    For i = 0 To n - 1
                        Set v = GetValue(buffer, embOffset)
                        id = v.value
                        Set v = GetValue(buffer, embOffset)
                        objOffset = v.value
                        embXRefTable.Add id, objOffset
                    Next i
                Else
                    Debug.Print "Missing count of embedded objects!"
                    If firstOffset >= 0 Then
                        Do
                            Set v = GetValue(buffer, embOffset)
                            id = v.value
                            Set v = GetValue(buffer, embOffset)
                            offset = v.value
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
                Set obj.value = GetValue(buffer, embOffset)
            Else
                Debug.Print "Error reading embedded object!"
                Stop
            End If
        Else
            Set obj = GetValue(content, offset)
        End If
    Else
NullValue:
        Set obj = pdfValue.NewValue(Null)
    End If
    
    Set loadObject = obj
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
        Set root = loadObject(content, xrefTable, root.value)
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


' updates objects Dictionary with all objects under root node, indexed by object id, i.e. loads a chunk of the PDF document
' Note: if root is an indirect Reference, then will return the obj it refers, else returns Nothing
Sub GetObjectsInTree(ByRef root As pdfValue, ByRef content() As Byte, ByRef xrefTable As Dictionary, ByRef objects As Dictionary)
    On Error GoTo errHandler
    Dim obj As pdfValue
    Dim v As Variant, k As Variant
    DoEvents
    'Debug.Print BytesToString(serialize(root))
    Select Case root.valueType
        Case PDF_ValueType.PDF_Boolean, PDF_ValueType.PDF_Comment, PDF_ValueType.PDF_Integer, PDF_ValueType.PDF_Name, PDF_ValueType.PDF_Null, PDF_ValueType.PDF_Real, PDF_ValueType.PDF_String
            ' Nothing to do
        Case PDF_ValueType.PDF_Array
            For Each v In root.value
                Set obj = v
                GetObjectsInTree obj, content, xrefTable, objects
            Next v
        Case PDF_ValueType.PDF_Dictionary
            For Each k In root.value.Keys
                Set obj = root.value.Item(k)
                GetObjectsInTree obj, content, xrefTable, objects
            Next k
        Case PDF_ValueType.PDF_Object
            GetObjectsInTree root.value, content, xrefTable, objects
        Case PDF_ValueType.PDF_Reference
            ' we need to load object
            If Not objects.Exists(CLng(root.value)) Then
                Set obj = loadObject(content, xrefTable, root.value)
                objects.Add CLng(root.value), obj
                GetObjectsInTree obj, content, xrefTable, objects
            End If
        Case PDF_ValueType.PDF_Stream
            Dim stream As pdfStream
            Set stream = root.value
            GetObjectsInTree stream.stream_meta, content, xrefTable, objects
            GetObjectsInTree stream.stream_data, content, xrefTable, objects
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


' reads in byteCount bytes and returns value as Long, offset is updated
Private Function getInt(ByRef content() As Byte, ByRef offset As Long, ByVal byteCount As Long, ByVal defaultValue As Long) As Long
    If byteCount < 1 Then
        getInt = defaultValue
    Else
        Dim i As Long, value As Long
        value = 0
        For i = 0 To byteCount - 1
            value = (value * 256) + content(offset)
            offset = offset + 1
        Next i
        getInt = value
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
        Set xrefStream = GetValue(content, offset) ' we don't keep these objects, so don't need to keep reference objs either
        If xrefStream.valueType <> PDF_ValueType.PDF_Object Then Exit Function ' wrong value returned, didn't find xref table
        Dim objStream As pdfStream
        Set objStream = xrefStream.value.value
        If Not objStream.Meta.Exists("/Type") Then Exit Function
        Dim obj As pdfValue
        Set obj = objStream.Meta.Item("/Type")
        If obj.valueType <> PDF_ValueType.PDF_Name Then Exit Function
        If Not IsMatch(obj.value, "/XRef") Then Exit Function
        
        ' TODO merge? or replacing from stream if both exist
        Set trailer = New pdfValue
        Set trailer = pdfValue.NewValue(pdfValue.NewValue(objStream.Meta, "/Dictionary"), "/Trailer")
        
        ' get Size value from trailer (really stream dictionary), # of expected total entries
        Size = GetXrefSize(trailer)
        
        ' see how many entries in this xref stream
        Dim subSections As Collection
        If objStream.Meta.Exists("/Index") Then
            Set subSections = objStream.Meta.Item("/Index").value
        Else ' use defaults
            Set subSections = New Collection
            subSections.Add pdfValue.NewValue(0)
            subSections.Add pdfValue.NewValue(Size)
        End If
        
        ' get the /W idth array for each entry
        If Not objStream.Meta.Exists("/W") Then Exit Function
        Dim widths(0 To 2) As Long
        Dim w As pdfValue
        Set w = objStream.Meta.Item("/W")
        widths(0) = CLng(w.value.Item(1).value)
        widths(1) = CLng(w.value.Item(2).value)
        widths(2) = CLng(w.value.Item(3).value)
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
            id = idObj.value
            Set idObj = subSections.Item(ndx + 2)
            Count = idObj.value
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
    Set prevXref = GetDictionaryValue(trailer.value, "/Prev")
    If prevXref.valueType <> PDF_ValueType.PDF_Null Then
        offset = prevXref.value
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

