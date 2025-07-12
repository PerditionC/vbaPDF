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

' optional named /Dest[ination]s
' this does not refer to an obj in the PDF but to a Dictionary key'd by /Dests in the /Root document catalog
' each key in the dictionary is the destination Name, and the value is either an array or dictionary with /D key having the array
' the array is a Destination, see PDF spec for all variations
Private m_Dests As pdfValue

' optional named /Outlines
Private m_Outlines As pdfValue


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

Public Enum PDF_FIT
    PDF_XYZ = -1
    PDF_FIT = 0
    PDF_FITH
    PDF_FITV
    PDF_FITR
    PDF_FITB
    PDF_FITBH
    PDF_FITBV
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
    If IsBlank(pdfTitle) Then
        If Info.hasKey("/Title") Then Info.asDictionary().Remove ("/Title")
    Else
        Info.asDictionary.item("/Title") = pdfTitle
    End If
End Property
Public Property Get Title() As String
    If IsBlank(m_Title) Then
        If Info.hasKey("/Title") Then
            m_Title = Info.asDictionary().item("/Title").value
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
                Set pdfRefObj = catalogDict.item("/Metadata")
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


' stores a cached copy of [optional] obj from /Root document catalog
' Note: assumes rootCatalog already established (loaded from file or otherwise Set)
Private Function objFromDocCatalog(ByRef m_obj As pdfValue, ByRef keyName As String, Optional ByVal isOptional As Boolean = True) As pdfValue
    On Error GoTo errHandler
    If m_obj Is Nothing Then Set m_obj = New pdfValue
    If m_obj.valueType = PDF_ValueType.PDF_Null Then
        ' find /Pages reference in /Root catalog and returns corresponding object from pdf or objectCache Dictionary
        If rootCatalog.hasKey(keyName) Then ' /key in the catalog dictionary << >>
            Dim obj As pdfValue
            Set obj = rootCatalog.asDictionary.item(keyName)
            ' obj should now be a reference to our /Pages object
            If obj.valueType = PDF_ValueType.PDF_Reference Then
                ' retrieve object (not reference) matching given obj reference's (id, generation) pair
                Set m_obj = getObject(CLng(obj.value), obj.generation)
            Else
                ' just return whatever we found
                ' if this is /Pages then probably an error as it should be a reference, but others may
                ' be a dictionary of values
                Set m_obj = obj
            End If
        Else
            If Not isOptional Then
                Debug.Print "PDF missing required " & keyName & " from /Root document catalog!"
                Stop
            End If
        End If
    End If
    Set objFromDocCatalog = m_obj
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function


' get/set Pages object
Public Property Set Pages(ByRef Pages As pdfValue)
    Set m_Pages = Pages
End Property
Public Property Get Pages() As pdfValue
    Set Pages = objFromDocCatalog(m_Pages, "/Pages", isOptional:=False) ' should return a PDF_ValueType=PDF_Object
End Property


' get/set optional named /Dest[ination]s
Public Property Set Dests(ByRef Dests As pdfValue)
    Set m_Dests = Dests
End Property
Public Property Get Dests() As pdfValue
    Set Dests = objFromDocCatalog(m_Dests, "/Dests") ' should return a PDF_ValueType=PDF_Dictionary
End Property


' get/set optional named /Outlines
Public Property Set Outlines(ByRef Outlines As pdfValue)
    Set m_Outlines = Outlines
End Property
Public Property Get Outlines() As pdfValue
    Set Outlines = objFromDocCatalog(m_Outlines, "/Outlines")
End Property


' returns /Count from /Pages (total /Page objects in pdf)
' readonly - add/remove /Page objects and update /Pages tree to alter value
Public Property Get pageCount() As Long
    If Pages.hasKey("/Count") Then
        pageCount = CLng(Pages.asDictionary().item("/Count").value)
    End If
End Property


' helper methods when creating a new pdf document or merging multiple ones

' adds & initializes the /Trailer/Info dictionary
' Note: if /Info already exists in /Trailer, this will replace it with a new one, assumes trailer is valid
Public Sub AddInfo(Optional defaults As Variant)
    Set Info = pdfValue.NewValue(pdfValue.NewValue(New Dictionary), id:=nextObjId) ' also updates /Trailer
    
    ' add in some additional default values
    With Info.asDictionary()
        Set .item("/Producer") = pdfValue.NewValue("vbaPDF")
    End With
End Sub


' adds & initializes the top level /Root document catalog
' Note: if /Root already exists in /Trailer, this will replace it with a new one, assumes trailer is valid
Public Sub NewDocumentCatalog()
    Set rootCatalog = pdfValue.NewValue(pdfValue.NewValue(New Dictionary), id:=nextObjId)
    ' add in some additional default values
    With rootCatalog.asDictionary()
        Set .item("/Type") = pdfValue.NewNameValue("/Catalog")
        Set .item("/Version") = pdfValue.NewNameValue("/" & Format(version / 10#, "#.0"))
        Set .item("/PageLayout") = pdfValue.NewNameValue("/OneColumn")  ' if not specified default is "/SinglePage"
        Set .item("/PageMode") = pdfValue.NewNameValue("/UseNone")      '  " is "/UseNone"
        Set .item("/Lang") = pdfValue.NewValue("en")                    '  " is unknown language
    End With
    
    updateTrailerReference "/Root", rootCatalog
End Sub


' creates a /typeName pdfValue dictionary object and updates /Root
Private Function NewTopLevelDictionary(ByRef typeName As String, ByRef parent As pdfValue, Optional ByRef defaults As Dictionary = Nothing) As pdfValue
    Dim newObj As pdfValue
    Set newObj = pdfValue.NewValue(pdfValue.NewValue(New Dictionary), id:=nextObjId)
    ' add in some default values
    With newObj.asDictionary()
        Set .item("/Type") = pdfValue.NewNameValue(typeName)
        ' /Parent should not exist in /Root/typeName, but need to update /Root to point to our new toplevel /typeName
        If parent Is Nothing Then
            Set rootCatalog.asDictionary(typeName) = newObj.referenceObj
        Else
            If parent.valueType = PDF_ValueType.PDF_Object Then
                Set .item("/Parent") = parent.referenceObj
            ElseIf parent.valueType = PDF_ValueType.PDF_Reference Then
                Set .item("/Parent") = parent ' already a reference
            Else
                Stop ' unsupported value!
            End If
        End If
        ' additional optional defaults
        If Not defaults Is Nothing Then
            Dim v As Variant
            For Each v In defaults.Keys
                Set .item(v) = defaults(v)
            Next
        End If
    End With
    
    ' ensure /typeName object is saved and easily retrievable
    objectCache.Add newObj.id, newObj
    Set NewTopLevelDictionary = newObj
End Function


' creates a default initialized /Pages object
Public Function NewPages(ByRef parent As pdfValue, Optional ByRef defaults As Dictionary = Nothing) As pdfValue
    If defaults Is Nothing Then Set defaults = New Dictionary
    With defaults
        Set .item("/Kids") = pdfValue.NewValue(New Collection)
        Set .item("/Count") = pdfValue.NewValue(0)
    End With
    Set NewPages = NewTopLevelDictionary("/Pages", parent, defaults)
End Function
' creates a default initialized /Page object
Public Function NewPage(ByRef parent As pdfValue, Optional ByRef defaults As Dictionary = Nothing) As pdfValue
    Set NewPage = NewTopLevelDictionary("/Page", parent, defaults)
End Function

' adds a /Page or /Pages to document
' if pages is Nothing then will also set top level /Root /Pages to newPages
' Note: currently all /Page or /Pages [thePages] are added to top level /Pages, in future may balance tree better (automatically add nested /Pages)
Public Sub AddPages(Optional ByRef thePages As pdfValue = Nothing)
    ' ensure top level /Pages is initialized, use m_Pages as we don't want to try loading /Root/Pages if creating new pdf doc
    If m_Pages.valueType = PDF_ValueType.PDF_Null Then
        Set Pages = NewPages(Nothing)
    End If
    
    ' add new /Pages or /Page to our top level /Pages
    If Not thePages Is Nothing Then
        ' first get Kids
        Dim kids As Collection
        ' add our Page(s) as a new child (as an indirect object reference)
        If Not Pages.asDictionary.Exists("/Kids") Then
            Pages.asDictionary.Add "/Kids", pdfValue.NewValue(New Collection)
        End If
        Set kids = Pages.asDictionary("/Kids").value
        kids.Add thePages.referenceObj
        Set thePages.asDictionary("/Parent") = Pages.referenceObj
        ' update page count
        Set Pages.asDictionary("/Count") = pdfValue.NewValue(pageCount + CLng(thePages.asDictionary("/Count").value))
    End If
End Sub


' creates a default initialized named /Dests object
Public Function NewDests(Optional defaults As Variant) As pdfValue
    Set NewDests = pdfValue.NewValue(New Dictionary)    ' << /Name1 [page Fit] /Name2 [page Fit] >> directly in /Root /Dests << ... >>
    Set rootCatalog.asDictionary("/Dests") = NewDests
End Function

' adds a /Dests dictionary to document catalog if needed
' then will add an entry with key destName and value theDest to the /Dests dictionary
' theDest should be a Destination, see NewDestination(...)
' destName should be a pdfValue of type PDF_Name
Public Sub AddNamedDestinations(ByRef destName As pdfValue, ByRef theDest As pdfValue)
    ' ensure top level /Dests is initialized
    If Dests.valueType = PDF_ValueType.PDF_Null Then
        Set Dests = NewDests
    End If
    
    ' add new Explicit Destination as Named value to our top level /Dests
    If Not ((theDest Is Nothing) Or IsBlank(destName.value)) Then
        Dim dict As Dictionary
        Set dict = Dests.asDictionary()
        If Not dict Is Nothing Then
            ' Note: the name for the destination is recommended to be a PDF_Name, but could just be a PDF_String
            Set dict(destName) = theDest
        Else
            Stop ' internal error or unexpected value, should have been a valid VBA Dictionary object
        End If
    End If
End Sub

' given an object that represents a Destination, parse into subparts (or given parts return new Destination object)
' if dest is a reference will attempt to Get the Object
' otherwise dest should be an array [page fit ...] or a dictionary with a "/D" key containing the array
' if it is a dictionary, the Dictionary object is returned in extra for any associated values (otherwise extra is Nothing)
' depending on fit, the other values will be corresponding values provided or -1 to indicated not provided / not applicable
Public Sub parseDestination(ByRef dest As pdfValue, ByRef page As Long, ByRef fit As PDF_FIT, ByRef leftX As Variant, ByRef rightX As Variant, ByRef topY As Variant, ByRef bottomY As Variant, ByRef zoom As Variant, ByRef extra As Dictionary)
    On Error GoTo errHandler
    page = 0: fit = PDF_FIT.PDF_FIT: leftX = Null: rightX = Null: topY = Null: bottomY = Null: zoom = 0: Set extra = Nothing
    
    Dim destArray As Collection
    If dest.valueType = PDF_ValueType.PDF_Array Then
        Set destArray = dest.value
    ElseIf dest.hasKey("/D") Then
        Set extra = dest.asDictionary()
        Set destArray = dest.asDictionary().item("/D").value
    End If
    
    page = CLng(destArray(0).value)
    Select Case TrimWS(destArray(1).value)
        Case "/XYZ"
            fit = PDF_FIT.PDF_XYZ
            leftX = CLng(destArray(2).value)
            topY = CLng(destArray(3).value)
            zoom = CLng(destArray(4).value)
        Case "/Fit"
            fit = PDF_FIT.PDF_FIT
        Case "/FitH"
            fit = PDF_FIT.PDF_FITH
            topY = CLng(destArray(2).value)
        Case "/FitV"
            fit = PDF_FIT.PDF_FITV
            leftX = CLng(destArray(2).value)
        Case "/FitR"
            fit = PDF_FIT.PDF_FITR
            leftX = CLng(destArray(2).value)
            bottomY = CLng(destArray(3).value)
            rightX = CLng(destArray(4).value)
            topY = CLng(destArray(5).value)
        Case "/FitB"
            fit = PDF_FIT.PDF_FITB
        Case "/FitBH"
            fit = PDF_FIT.PDF_FITBH
            topY = CLng(destArray(2).value)
        Case "/FitBV"
            fit = PDF_FIT.PDF_FITBV
            leftX = CLng(destArray(2).value)
        Case Else
            Stop ' unknown option
    End Select
    
    Exit Sub
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Sub
Public Function NewDestination(ByRef page As Long, Optional ByRef fit As PDF_FIT = PDF_FIT.PDF_FIT, Optional ByRef leftX As Variant = Null, Optional ByRef rightX As Variant = Null, Optional ByRef topY As Variant = Null, Optional ByRef bottomY As Variant = Null, Optional ByRef zoom As Variant = Null, Optional ByRef extra As Dictionary = Nothing) As pdfValue
    On Error GoTo errHandler
    Dim dest As pdfValue
    Dim destArray As Collection: Set destArray = New Collection
    If extra Is Nothing Then
        Set dest = pdfValue.NewValue(destArray)
    Else
        Set extra.item("/D") = pdfValue.NewValue(destArray)
        Set dest = pdfValue.NewValue(extra)
    End If
    
    destArray.Add pdfValue.NewValue(page)
    Select Case fit
        Case PDF_FIT.PDF_XYZ
            destArray.Add pdfValue.NewNameValue("/XYZ")
            destArray.Add pdfValue.NewValue(leftX)
            destArray.Add pdfValue.NewValue(topY)
            destArray.Add pdfValue.NewValue(zoom)
        Case PDF_FIT.PDF_FIT
            destArray.Add pdfValue.NewNameValue("/Fit")
        Case PDF_FIT.PDF_FITH
            destArray.Add pdfValue.NewNameValue("/FitH")
            destArray.Add pdfValue.NewValue(topY)
        Case PDF_FIT.PDF_FITV
            destArray.Add pdfValue.NewNameValue("/FitV")
            destArray.Add pdfValue.NewValue(leftX)
        Case PDF_FIT.PDF_FITR
            destArray.Add pdfValue.NewNameValue("/FitR")
            destArray.Add pdfValue.NewValue(leftX)
            destArray.Add pdfValue.NewValue(bottomY)
            destArray.Add pdfValue.NewValue(rightX)
            destArray.Add pdfValue.NewValue(topY)
        Case "FITB"
            destArray.Add pdfValue.NewNameValue("/FitB")
        Case PDF_FIT.PDF_FITBH
            destArray.Add pdfValue.NewNameValue("/FitBH")
            destArray.Add pdfValue.NewValue(topY)
        Case PDF_FIT.PDF_FITBV
            destArray.Add pdfValue.NewNameValue("/FitBV")
            destArray.Add pdfValue.NewValue(leftX)
        Case Else
            Stop ' unknown option
    End Select
    
    Set NewDestination = dest
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function


' creates a default initialized /Outlines object
' defaults should include /First /Last and /Count
Public Function NewOutlines(ByRef parent As pdfValue, Optional ByRef defaults As Dictionary = Nothing) As pdfValue
    If defaults Is Nothing Then Set defaults = New Dictionary
    With defaults
        'If Not .Exists("/First") Then Set .item("/First") = pdfValue.NewValue(Null)
        'If Not .Exists("/Last") Then Set .item("/Last") = pdfValue.NewValue(Null)
        If Not .Exists("/Count") Then Set .item("/Count") = pdfValue.NewValue(0)
    End With
    Set NewOutlines = NewTopLevelDictionary("/Outlines", parent, defaults)
End Function

' add (or sets) top level /Root /Outlines of document
' if anOutlineItem is Nothing then creates default toplevel /Outlines item if doesn't already exist
' otherwise sets Outlines to specified object
Public Sub AddOutlines(Optional ByRef anOutlineItem As pdfValue = Nothing)
    If anOutlineItem Is Nothing Then
        If Outlines.valueType = PDF_ValueType.PDF_Null Then
            Set Outlines = NewOutlines(Nothing)
            ' don't forget to add to our object cache
            Set objectCache(Outlines.ID) = Outlines
        End If
    Else
        Set Outlines = anOutlineItem
    End If
End Sub

' follows a linked list and returns pdfValue at end
' item can be obj or obj reference, returns obj reference
Private Function GetEndOfChain(ByRef item As pdfValue, ByRef linkName As String) As pdfValue
    Dim nullValue As pdfValue: Set nullValue = pdfValue.NewValue(Null)
    Dim obj As pdfValue
    Dim nextItem As pdfValue: Set nextItem = item
    If nextItem Is Nothing Then Set nextItem = nullValue
    ' we expect item and nextItem to object references, but we need actual obj to follow linked list
    If nextItem.valueType = PDF_ValueType.PDF_Reference Then Set obj = getObject(nextItem.value, nextItem.generation)
    Set GetEndOfChain = nextItem.referenceObj
    Do While nextItem.valueType = PDF_ValueType.PDF_Dictionary
        Dim dict As Dictionary: Set dict = nextItem.asDictionary()
        Set nextItem = nullValue
        If dict.Exists(linkName) Then
            Set nextItem = dict.item(linkName)
            If nextItem.valueType = PDF_ValueType.PDF_Reference Then Set obj = getObject(nextItem.value, nextItem.generation)
            If nextItem.valueType = PDF_ValueType.PDF_Dictionary Then
                Set GetEndOfChain = nextItem.referenceObj
            End If
        End If
    Loop
End Function


' creates a default initialized /Outline item object
' defaults should include /First /Last along with /Count, /Title /Prev /Next and optionally /Dest
Public Function NewOutlineItem(ByRef parent As pdfValue, Optional ByRef defaults As Dictionary) As pdfValue
    If defaults Is Nothing Then Set defaults = New Dictionary
    With defaults
        'If Not .Exists("/Next") Then Set .Item("/Next") = pdfValue.NewValue(Null)
        'If Not .Exists("/Prev") Then Set .Item("/Prev") = pdfValue.NewValue(Null)
        'If Not .Exists("/Count") Then Set .item("/Count") = pdfValue.NewValue(1)
    End With
    ' Note: /Type /OutlineItem doesn't exist in spec, but should be safe as no /Type specified & additional entries allowed
    Set NewOutlineItem = NewTopLevelDictionary("/OutlineItem", parent, defaults)
End Function

' adds a /Outlines to document, specifies parent /Outline
' if parent is Nothing or PDF_Null then will use root /Outline (.Outlines)
' if Outlines is Nothing then will also set top level /Root /Outlines to newOutlines
'
' If OutlineItem does not have its /Next and /Prev values set already then adds to end of list
' we update it's /Parent and if needed /First and /Last of /Outline
' OutlineItem's /Count is added to Outline's /Count
Public Sub AddOutlineItem(ByRef parent As pdfValue, Optional ByRef anOutlineItem As pdfValue = Nothing)
    ' ensure top level dictionary is initialized
    If parent Is Nothing Then Set parent = pdfValue.NewValue(Null)
    If Outlines.valueType = PDF_ValueType.PDF_Null Then
        Set Outlines = NewOutlines(Nothing)
    End If
    If parent.valueType = PDF_ValueType.PDF_Null Then Set parent = Outlines
    
    ' insert a new /OutlineItem into our outlines
    If Not anOutlineItem Is Nothing Then
        ' ensure parent is correctly set
        Set anOutlineItem.asDictionary("/Parent") = parent.referenceObj
        ' see if we need to assign and id and add to our objectcache
        If anOutlineItem.ID < 0 Then
            anOutlineItem.ID = nextObjId
            anOutlineItem.generation = 0
            objectCache.Add anOutlineItem.ID, anOutlineItem
        End If
    
        ' determine if /First and /Last exists or need to be created
        ' determine if this item is being appended to end or has picked its place
        Dim first As pdfValue, last As pdfValue, prevItem As pdfValue, nextItem As pdfValue
        With parent.asDictionary()
            If .Exists("/First") Then
                Set first = .item("/First")
            Else
                Set first = pdfValue.NewValue(Null)
            End If
            If .Exists("/Last") Then
                Set last = .item("/Last")
            Else
                Set last = pdfValue.NewValue(Null)
            End If
            
            If anOutlineItem.asDictionary().Exists("/Prev") Then
                Set prevItem = anOutlineItem.asDictionary().item("/Prev")
            Else
                Set prevItem = pdfValue.NewValue(Null)
            End If
            If anOutlineItem.asDictionary().Exists("/Next") Then
                Set nextItem = anOutlineItem.asDictionary().item("/Next")
            Else
                Set nextItem = pdfValue.NewValue(Null)
            End If
            
            If (first.valueType = PDF_ValueType.PDF_Null) Then
                If (prevItem.valueType = PDF_ValueType.PDF_Null) Then
                    Set .item("/First") = anOutlineItem.referenceObj
                Else
                    Set .item("/First") = GetEndOfChain(prevItem, "/Prev")   ' returns referenceObj
                End If
            Else
                If (prevItem.valueType = PDF_ValueType.PDF_Null) And (last.valueType <> PDF_ValueType.PDF_Null) Then
                    Set anOutlineItem.asDictionary().item("/Prev") = last
                    Dim obj As pdfValue
                    Set obj = getObject(last.value, last.generation)
                    Set obj.asDictionary().item("/Next") = anOutlineItem.referenceObj
                    Set .item("/Last") = anOutlineItem.referenceObj
                'Else leave /Prev whatever it is
                End If
            End If
            If (last.valueType = PDF_ValueType.PDF_Null) Then
                If (nextItem.valueType = PDF_ValueType.PDF_Null) Then
                    Set .item("/Last") = anOutlineItem.referenceObj
                Else
                    Set .item("/Last") = GetEndOfChain(nextItem, "/Next")    ' returns referenceObj
                End If
            Else
                If (nextItem.valueType = PDF_ValueType.PDF_Null) And (first.valueType <> PDF_ValueType.PDF_Null) Then
                    'Set anOutlineItem.asDictionary().item("/Next") = first
                'Else leave /Next as is
                End If
            End If
            
            Dim count As Long
            Dim countValue As pdfValue
            If anOutlineItem.hasKey("/Count") Then
                count = anOutlineItem.asDictionary().item("/Count").value
            Else
                count = 1
            End If
            If .Exists("/Count") Then
                count = count + .item("/Count").value
            End If
            Set .item("/Count") = pdfValue.NewValue(count)
        End With
    End If
End Sub



' updates id of all objects under root pdfValue beginning at baseId
' Note: if root is Nothing then will update whole document including caches
' Warning: assumes each obj id refers to same obj and objReference instances
Public Function renumberIds(ByVal baseId As Long, Optional ByRef root As pdfValue = Nothing, Optional ByRef visited As Dictionary = Nothing) As Long
    On Error GoTo errHandler
    Dim v As Variant
    Dim obj As pdfValue
    Dim fixCache As Boolean
    
    ' renumber whole tree?
    If root Is Nothing Then
        fixCache = True
        Set root = trailer
    End If
    
    ' should be Nothing on first call, we use this to avoid recursion and thus renumber already renumbered objs
    If visited Is Nothing Then
        Set visited = New Dictionary
    End If
    
    ' don't renumber same object more than once
    If root.id <> 0 Then
        If visited.Exists(root) Then
            renumberIds = baseId
            Exit Function
        End If
        ' and actually do the renumbering
        root.id = baseId
        baseId = baseId + 1
        ' Note: we use actual object and not id here as temp there may be multiple objs with same id
        ' e.g. obj 1 renumbered to 4, and there is still an object 4 not yet renumbered
        visited.Add root, root
    End If
    ' update any children
    Select Case root.valueType
        Case PDF_ValueType.PDF_Array
            For Each v In root.value
                Set obj = v
                baseId = renumberIds(baseId, obj, visited)
            Next v
        Case PDF_ValueType.PDF_Dictionary
            For Each v In root.value.Items
                Set obj = v
                baseId = renumberIds(baseId, obj, visited)
            Next v
        Case PDF_ValueType.PDF_Reference
            Set obj = getObject(root.value, 0)
            baseId = renumberIds(baseId, obj, visited)
        Case PDF_ValueType.PDF_Trailer, PDF_ValueType.PDF_Object
            baseId = renumberIds(baseId, root.value, visited)
        Case Else
            ' no nested objects that need id's updated
    End Select
    ' last id used, or same as provided if nothing updated
    renumberIds = baseId
    
    ' because we renumbered everything, we need to invalidate or rebuild our cache's
    ' but only from top level call
    ' note: because we have renumbered, we also invalidate our content buffer as it refers to old ids
    If fixCache Then
        'Set visited = Nothing ' this is ByRef so only set to Nothing if originally Nothing
        
        If ByteArraySize(content) >= 0 Then Erase content
        Dim newDict As Dictionary
        If Not objectCache Is Nothing Then
            Set newDict = New Dictionary
            For Each v In objectCache.Items
                Set obj = v
                If obj.id <> 0 Then Set newDict(obj.id) = obj
            Next v
            Set objectCache = newDict
            Set newDict = Nothing
        End If
    
        If Not refCache Is Nothing Then
            Set newDict = New Dictionary
            For Each v In refCache.Items
                Set obj = v
                With obj
                    If .id <> 0 Then
                        Set newDict(.id & "_" & .generation) = obj
                        Set newDict(.id & " " & .generation & " R") = obj.referenceObj
                    End If
                End With
            Next v
            Set refCache = newDict
            Set newDict = Nothing
        End If
    
        Set sosCache = Nothing
    End If
    
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
        Set obj = loadObject(id)
        If obj Is Nothing Then Set obj = New pdfValue
        
        ' if successfully loaded, update our object cache
        If cacheObject And (obj.valueType <> PDF_ValueType.PDF_Null) Then
            If objectCache.Exists(id) Then
                ' we need to determine if same generation, presumably not or we would have retrieved it above
                Dim dict As Dictionary
                Dim curObj As pdfValue
                Dim v As Variant
                Set v = objectCache.item(id)
                If typeName(v) = "pdfValue" Then
                    Set curObj = objectCache.item(id)
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
                ElseIf typeName(v) = "Dictionary" Then
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
        Set v = objectCache.item(id)
        If typeName(v) = "pdfValue" Then
            Set getCachedObject = v
            If getCachedObject.generation = generation Then Exit Function
        ElseIf typeName(v) = "Dictionary" Then
            ' assume may contain multiple generations, key'd by generation
            If v.Exists(generation) Then
                Set getCachedObject = v.item(generation)
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
    
    If (obj.valueType = PDF_ValueType.PDF_Stream) Then
        ' see if its a stream object stream (any other object stream  written out asis)
        If obj.asDictionary.Exists("/Type") Then
            Dim pdfNameValue As pdfValue
            Set pdfNameValue = obj.asDictionary.item("/Type")
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

    
' save all nested /Outlines objects beginning with outlineParentObj
' pass in .Outlines to save full outline hierarchy, or any subtree as needed
Sub SaveOutlinePdfObjects(ByRef outputFileNum As Integer, ByRef outlineParentObj As pdfValue, _
            ByRef offset As Long, Optional ByVal baseId As Long = 0, Optional ByVal prettyPrint As Boolean = True)
    On Error GoTo ErrHandler
    SavePdfObject outputFileNum, outlineParentObj, offset, baseId, prettyPrint
    If outlineParentObj.hasKey("/First") Then
        Dim outlineObj As pdfValue
        Set outlineObj = outlineParentObj.asDictionary().item("/First")
        ' next line should always be true, only an indirect reference not object itself
        If outlineObj.valueType = PDF_ValueType.PDF_Reference Then Set outlineObj = getObject(outlineObj.value, outlineObj.generation)
        Do While Not outlineObj Is Nothing
            ' recursively save outline tree
            SaveOutlinePdfObjects outputFileNum, outlineObj, offset, baseId, prettyPrint
            If outlineObj.hasKey("/Next") Then
                Set outlineObj = outlineObj.asDictionary().item("/Next")
                Set outlineObj = getObject(outlineObj.value, outlineObj.generation)
            Else
                Set outlineObj = Nothing
            End If
        Loop
    End If
    Exit Sub
ErrHandler:
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
    maxId = xrefTable.count - 1 ' should be at least this high
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
            Set entry = xrefTable.item(ndx)
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
        If IsEmpty(xrefTable.item(id)) Then
            Stop
            GoTo newEntry
        End If
        ' update existing entry
        Set entry = xrefTable.item(id)
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
    Set Pages = New pdfValue ' default is null
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
    Set Pages = Nothing
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
    Dim token As String: token = chr(bytes(offset))
    
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
            If IsMatch("<", chr(bytes(offset + 1))) Then
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
' from spec all Names should be UTF-8 encoded, so always convert from UTF-8 (BOM or not)
Public Function ProcessName(ByRef name() As Byte) As String
    Dim buffer() As Byte
    ReDim buffer(0 To UBound(name))
    Dim nameNdx As Long, bufNdx As Long, val As Long
    For nameNdx = LBound(name) To UBound(name)
        val = name(nameNdx)
        If val = 35 Then ' #
            Dim s As String
            s = "&H" & chr(name(nameNdx + 1)) & chr(name(nameNdx + 2))  ' #00 --> &H00
            nameNdx = nameNdx + 2 ' skip past 00 also
            buffer(bufNdx) = CLng(s)
        Else
            buffer(bufNdx) = val
        End If
        bufNdx = bufNdx + 1
    Next nameNdx
    ReDim Preserve buffer(0 To bufNdx - 1) ' shrink to size (or no change if no #00 encoded values)
    ' if UTF8 then convert to VBA Unicode String, note / is not considered part of Name so skip past it looking for BOM
    If HasUtf8BOM(buffer, 1) Then
        ProcessName = "/" & Mid(Utf8BytesToString(buffer), 3) ' skip past BOM [/ BOM Name]
    Else
        ProcessName = Utf8BytesToString(buffer)
    End If
End Function
' opposite of ProcessName, converts spaces and other weird characters in #00 form
' UTF8BOM should only be added for custom names, all standard names are ASCII and should not include BOM
' from spec all Names should be UTF-8 encoded, so no need for a BOM
Public Function EscapeName(ByRef name As String, Optional ByVal addUtf8BOM As Boolean = False) As String
    Dim ndx As Long, chrA As Long
    Dim c As String
    Dim s As String
    If Left(name, 1) = "/" Then  ' / is not part of the Name, but we must ensure it is the start of it
        s = Mid(name, 2) ' skip /
    Else
        s = name
    End If
    ' copy over each character, but ensuring control codes <=32 & non ASCII >= 128 are encoded as hex digits
    ' we assume a String in VBA is a Unicode String so convert name to UTF8 bytes
    Dim utf8Name() As Byte: utf8Name = StringToUtf8Bytes(s, BOM:=addUtf8BOM)
    For ndx = LBound(utf8Name) To UBound(utf8Name)
        chrA = utf8Name(ndx)
        c = chr(chrA)
        If (chrA <= 32) Or (chrA >= 127) Then  ' from spec < Asc("!") or > Asc("~") need encoding
            EscapeName = EscapeName + "#" + Hex(chrA \ 16) + Hex(chrA Mod 16)
        Else
            EscapeName = EscapeName + c
        End If
    Next
    EscapeName = "/" & EscapeName
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
    Dim token As String: token = chr(bytes(offset))
    
    Dim tmpStr As String
    Dim name As pdfValue, value As pdfValue
    Select Case result.valueType
        Case PDF_ValueType.PDF_Null
            ' result.value = Empty
            If Not IsMatch(GetWord(bytes, offset), "null") Then
                Stop ' error, expecting "null"
            End If
        Case PDF_ValueType.PDF_Name
            Dim word() As Byte
            Dim count As Long
            GetAsBytes bytes, offset, word, count
            result.value = ProcessName(word)
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
                                Dim octStr As String: octStr = chr(bytes(offset))
                                Dim octVal As Long: octVal = 0
                                Dim maxOctStrLen As Integer: maxOctStrLen = 3
                                Do While (octStr >= "0") And (octStr <= "9") And (maxOctStrLen > 0)
                                    maxOctStrLen = maxOctStrLen - 1 ' only 1 to 3 octal digits
                                    octVal = (octVal * 8) + CLng(octStr)
                                    offset = offset + 1
                                    octStr = chr(bytes(offset))
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
                    HexStr = chr(bytes(offset))
                    offset = offset + 1
                    offset = SkipWhiteSpace(bytes, offset)
                    If bytes(offset) <> 62 Then ' Asc(">")
                        HexStr = HexStr & chr(bytes(offset))
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
                    tmpStr = tmpStr & chr(strBuffer(ndx2))
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
            If chr(bytes(offset)) = vbCr Then offset = offset + 1
            If chr(bytes(offset)) = vbLf Then offset = offset + 1
            result.valueType = PDF_ValueType.PDF_StreamData
            Dim dataBytes() As Byte
            If Meta Is Nothing Then Set Meta = New Dictionary
            If Meta.Exists("/Length") Then
                Dim byteLen As Long
                If IsObject(Meta.item("/Length").value) Then
                    ' Reference to actual /Length value, allowed but not common
                    Stop
                    ' TODO pass we need Dictionary of loaded objects or xrefTable to load value
                Else
                    ' /Length directly in file as number, usual
                    byteLen = CLng(Meta.item("/Length").value)
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
                    If chr(bytes(offset)) = vbCr Then
                        tmpStr = tmpStr & vbCr
                        offset = offset + 1
                    End If
                    If chr(bytes(offset)) = vbLf Then
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
' uses xrefTable to determine offset directly or obtain using cross reference stream
' due to potential slowness uncompressing in VBA, stream object streams should be cached
' the sosCache is only used for stream object streams and only if provided
' the refCache is to ensure only 1 indirect reference object instance is created for each object
' this will use provided VBA reference object instance instead of creating a new [different] one
' forceReload bypasses refCache and returns a new VBA object instance ***
' *** changes to this object's id may not effect any other object nor their reference values
' *** if object previously loaded returns a new instance, otherwise results same as if is False
' *** possibly useful to duplicate an object to be mutated and added, but use with care
Function loadObject(ByVal Index As Long, Optional forceReload As Boolean = False) As pdfValue
    On Error GoTo errHandler
    Dim obj As pdfValue
    
    ' speed up, ensure single object returned, and avoid issues when objects are removed from current objectCache
    If (Not forceReload) And (Not refCache Is Nothing) Then
        If refCache.Exists(Index) Then
            Set loadObject = refCache(Index)
            Exit Function
        End If
    End If
    
    If xrefTable.Exists(Index) Then
        Dim entry As xrefEntry
        Set entry = xrefTable.item(Index)
        If entry.isFree Or ((Not entry.isEmbeded) And (entry.offset <= 0)) Then GoTo nullValue
        Dim offset As Long: offset = entry.offset
        If offset >= ByteArraySize(content) Then GoTo nullValue
        If entry.isEmbeded Then
            Dim cntrObjEntry As xrefEntry
            Set cntrObjEntry = xrefTable.item(entry.embedObjId)
            Dim cntrObj As pdfValue
            ' try loading containing object (stream object stream) from cache before potentially uncompressing
            If Not sosCache Is Nothing Then
                If sosCache.Exists(entry.embedObjId) Then Set cntrObj = sosCache(entry.embedObjId)
            End If
            If cntrObj Is Nothing Then  ' not in cache or no cache provided
                Set cntrObj = loadObject(entry.embedObjId)
                If Not sosCache Is Nothing Then Set sosCache(entry.embedObjId) = cntrObj    ' add/update cache
            End If
            
            ' extract our embedded object
            If cntrObj.value.valueType <> PDF_ValueType.PDF_Stream Then
                Debug.Print "Error! expecting stream object stream!"
                Stop
                GoTo nullValue
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
                    firstOffset = CLng(dict.item("/First").value)
                Else
                    firstOffset = -1
                End If
                Dim n As Long, id As Long, objOffset As Long
                Dim v As pdfValue
                If dict.Exists("/N") Then
                    n = CLng(dict.item("/N").value)
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
                embOffset = embXRefTable.item(entry.id) + firstOffset
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
nullValue:
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
        Set root = loadObject(root.value)
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
                Set obj = root.value.item(k)
                GetObjectsInTree obj, content, xrefTable, objects
            Next k
        Case PDF_ValueType.PDF_Object
            GetObjectsInTree root.value, content, xrefTable, objects
        Case PDF_ValueType.PDF_Reference
            ' we need to load object
            If Not objects.Exists(CLng(root.value)) Then
                Set obj = loadObject(root.value)
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
    Dim count As Long   ' how many entries in current xref chunk
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
            count = CLng(str)
            Dim entryCount As Long: entryCount = count ' # of entries in this section
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
            Size = Size - count
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
        Set obj = objStream.Meta.item("/Type")
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
            Set subSections = objStream.Meta.item("/Index").value
        Else ' use defaults
            Set subSections = New Collection
            subSections.Add pdfValue.NewValue(0)
            subSections.Add pdfValue.NewValue(Size)
        End If
        
        ' get the /W idth array for each entry
        If Not objStream.Meta.Exists("/W") Then Exit Function
        Dim widths(0 To 2) As Long
        Dim w As pdfValue
        Set w = objStream.Meta.item("/W")
        widths(0) = CLng(w.value.item(1).value)
        widths(1) = CLng(w.value.item(2).value)
        widths(2) = CLng(w.value.item(3).value)
        Set w = Nothing
        
        ' we need the uncompressed (un-/Filter'd) data
        Dim rawData() As Byte
        rawData = objStream.udata
        
        Dim objOffset As Long
        objOffset = 0
        Dim ndx As Long
        Dim idObj As pdfValue
        For ndx = 0 To subSections.count - 1 Step 2
            ' get starting index (default value) and count of entries
            Set idObj = subSections.item(ndx + 1)
            id = idObj.value
            Set idObj = subSections.item(ndx + 2)
            count = idObj.value
            Set idObj = Nothing
            
            ' read in object id and offset information
            Dim i As Long
            For i = 0 To count - 1
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

