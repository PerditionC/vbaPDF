# pdfDocument.cls API Documentation

## Public Enumerations

### PDF_VERSIONS
Defines supported PDF version constants.

```vba
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
```

### PDF_FIT
Defines page fit modes for destinations and viewing.

```vba
Public Enum PDF_FIT
    PDF_XYZ = -1    ' Fit to specific coordinates with zoom
    PDF_FIT = 0     ' Fit entire page in window
    PDF_FITH        ' Fit page width in window
    PDF_FITV        ' Fit page height in window
    PDF_FITR        ' Fit rectangle in window
    PDF_FITB        ' Fit bounding box in window
    PDF_FITBH       ' Fit bounding box width in window
    PDF_FITBV       ' Fit bounding box height in window
End Enum
```

## Public API

### Document Management

#### pdfDocument Function
Creates a new pdfDocument instance with optional file loading.

```vba
Public Function pdfDocument(Optional ByVal filename As String = vbNullString) As pdfDocument
```

**Parameters:**
- `filename` (String, Optional): Path to PDF file to load. If empty, creates blank document.

**Returns:** 
- `pdfDocument`: New document instance, or Nothing if loading fails.

**Usage:**
```vba
' Create new blank document
Dim doc As pdfDocument
Set doc = pdfDocument.pdfDocument()

' Load existing PDF
Set doc = pdfDocument.pdfDocument("C:\path\to\file.pdf")
If doc Is Nothing Then
    Debug.Print "Failed to load PDF"
End If
```

#### loadPdf Function
Loads a PDF document from file.

```vba
Public Function loadPdf(ByVal pdfFilename As String) As Boolean
```

**Parameters:**
- `pdfFilename` (String): Full path to PDF file to load.

**Returns:**
- `Boolean`: True on success, False on error.

**Usage:**
```vba
Dim doc As New pdfDocument
If doc.loadPdf("C:\document.pdf") Then
    Debug.Print "PDF loaded successfully"
Else
    Debug.Print "Failed to load PDF"
End If
```

#### parsePdf Function
Parses the loaded PDF document and populates object cache.

```vba
Public Function parsePdf() As Boolean
```

**Returns:**
- `Boolean`: True on success, False on error.

**Usage:**
```vba
Dim doc As New pdfDocument
If doc.loadPdf("document.pdf") And doc.parsePdf() Then
    Debug.Print "PDF parsed successfully"
    Debug.Print "Page count: " & doc.pageCount
End If
```

#### savePdf Function
Saves the document to its original filename.

```vba
Public Function savePdf() As Boolean
```

**Returns:**
- `Boolean`: True on success, False on error.

**Usage:**
```vba
doc.Title = "Modified Document"
If doc.savePdf() Then
    Debug.Print "Document saved"
End If
```

#### savePdfAs Function
Saves the document to a new file.

```vba
Public Function savePdfAs(ByVal fileNameAndPath As String) As Boolean
```

**Parameters:**
- `fileNameAndPath` (String): Full path for new PDF file.

**Returns:**
- `Boolean`: True on success, False on error.

**Usage:**
```vba
If doc.savePdfAs("C:\output\modified.pdf") Then
    Debug.Print "Document saved as new file"
End If
```

### Document Properties

#### version Property
Gets or sets the PDF version.

```vba
Public Property Let version(ByVal pdfVersion As PDF_VERSIONS)
Public Property Get version() As PDF_VERSIONS
```

**Parameters:**
- `pdfVersion` (PDF_VERSIONS): Version to set.

**Returns:**
- `PDF_VERSIONS`: Current PDF version.

**Usage:**
```vba
doc.version = PDF_VERSIONS.PDF_1_7
Debug.Print "PDF Version: " & doc.version
```

#### Header Property
Gets or sets the PDF header string.

```vba
Public Property Let Header(ByVal pdfHeader As String)
Public Property Get Header() As String
```

**Parameters:**
- `pdfHeader` (String): Header string to set (e.g., "%PDF-1.7").

**Returns:**
- `String`: Current PDF header.

#### Title Property
Gets or sets the document title.

```vba
Public Property Let Title(pdfTitle As String)
Public Property Get Title() As String
```

**Parameters:**
- `pdfTitle` (String): Title to set. Empty string removes title.

**Returns:**
- `String`: Current document title or filename if no title set.

**Usage:**
```vba
doc.Title = "My Document"
Debug.Print "Title: " & doc.Title
```

#### pageCount Property (Read-Only)
Gets the total number of pages in the document.

```vba
Public Property Get pageCount() As Long
```

**Returns:**
- `Long`: Number of pages in document.

**Usage:**
```vba
Debug.Print "Document has " & doc.pageCount & " pages"
```

#### nextObjId Property
Gets or sets the next available object ID.

```vba
Public Property Let nextObjId(ByVal nextId As Long)
Public Property Get nextObjId() As Long
```

**Parameters:**
- `nextId` (Long): Next ID to use.

**Returns:**
- `Long`: Next available object ID (increments on each call).

### Document Structure Access

#### Info Property
Gets or sets the document information dictionary.

```vba
Public Property Set Info(ByRef Info As pdfValue)
Public Property Get Info() As pdfValue
```

**Parameters:**
- `Info` (pdfValue): Information dictionary to set.

**Returns:**
- `pdfValue`: Current document information dictionary.

**Usage:**
```vba
Dim info As pdfValue
Set info = doc.Info
If info.hasKey("/Author") Then
    Debug.Print "Author: " & info.asDictionary()("/Author").value
End If
```

#### Meta Property
Gets or sets the document metadata stream.

```vba
Public Property Set Meta(ByRef Meta As pdfValue)
Public Property Get Meta() As pdfValue
```

**Parameters:**
- `Meta` (pdfValue): Metadata stream to set.

**Returns:**
- `pdfValue`: Current document metadata stream.

#### Pages Property
Gets or sets the document pages tree.

```vba
Public Property Set Pages(ByRef Pages As pdfValue)
Public Property Get Pages() As pdfValue
```

**Parameters:**
- `Pages` (pdfValue): Pages tree to set.

**Returns:**
- `pdfValue`: Current pages tree object.

#### Dests Property
Gets or sets named destinations dictionary.

```vba
Public Property Set Dests(ByRef Dests As pdfValue)
Public Property Get Dests() As pdfValue
```

**Parameters:**
- `Dests` (pdfValue): Named destinations dictionary to set.

**Returns:**
- `pdfValue`: Current named destinations dictionary.

#### Outlines Property
Gets or sets document outline (bookmarks) tree.

```vba
Public Property Set Outlines(ByRef Outlines As pdfValue)
Public Property Get Outlines() As pdfValue
```

**Parameters:**
- `Outlines` (pdfValue): Outline tree to set.

**Returns:**
- `pdfValue`: Current outline tree object.

### Document Creation and Modification

#### AddInfo Subroutine
Adds and initializes the document information dictionary.

```vba
Public Sub AddInfo(Optional defaults As Variant)
```

**Parameters:**
- `defaults` (Variant, Optional): Additional default values to include.

**Usage:**
```vba
doc.AddInfo
doc.Info.asDictionary()("/Author") = "John Doe"
doc.Info.asDictionary()("/Subject") = "Test Document"
```

#### NewDocumentCatalog Subroutine
Creates and initializes the document catalog (root object).

```vba
Public Sub NewDocumentCatalog()
```

**Usage:**
```vba
Dim doc As New pdfDocument
doc.NewDocumentCatalog
```

#### NewPages Function
Creates a new pages tree object.

```vba
Public Function NewPages(ByRef parent As pdfValue, Optional ByRef defaults As Dictionary = Nothing) As pdfValue
```

**Parameters:**
- `parent` (pdfValue): Parent object reference.
- `defaults` (Dictionary, Optional): Additional default values.

**Returns:**
- `pdfValue`: New pages tree object.

#### NewPage Function
Creates a new page object.

```vba
Public Function NewPage(ByRef parent As pdfValue, Optional ByRef defaults As Dictionary = Nothing) As pdfValue
```

**Parameters:**
- `parent` (pdfValue): Parent pages tree reference.
- `defaults` (Dictionary, Optional): Additional default values.

**Returns:**
- `pdfValue`: New page object.

**Usage:**
```vba
Dim pagesTree As pdfValue
Set pagesTree = doc.NewPages(Nothing)
Dim newPage As pdfValue
Set newPage = doc.NewPage(pagesTree)
```

#### AddPages Subroutine
Adds pages to the document.

```vba
Public Sub AddPages(Optional ByRef thePages As pdfValue = Nothing)
```

**Parameters:**
- `thePages` (pdfValue, Optional): Pages object to add. If Nothing, initializes top-level pages tree.

**Usage:**
```vba
' Initialize pages tree
doc.AddPages

' Add a new page
Dim newPage As pdfValue
Set newPage = doc.NewPage(doc.Pages)
doc.AddPages newPage
```

### Named Destinations

#### NewDests Function
Creates a new named destinations dictionary.

```vba
Public Function NewDests(Optional defaults As Variant) As pdfValue
```

**Parameters:**
- `defaults` (Variant, Optional): Additional default values.

**Returns:**
- `pdfValue`: New destinations dictionary.

#### AddNamedDestinations Subroutine
Adds a named destination to the document.

```vba
Public Sub AddNamedDestinations(ByRef destName As pdfValue, ByRef theDest As pdfValue)
```

**Parameters:**
- `destName` (pdfValue): Name of the destination (PDF_Name type).
- `theDest` (pdfValue): Destination array or dictionary.

**Usage:**
```vba
Dim destName As pdfValue
Set destName = pdfValue.NewNameValue("/MyDest")
Dim dest As pdfValue
Set dest = doc.NewDestination(1, PDF_FIT.PDF_FIT)
doc.AddNamedDestinations destName, dest
```

#### NewDestination Function
Creates a new destination object.

```vba
Public Function NewDestination(ByRef page As Long, Optional ByRef fit As PDF_FIT = PDF_FIT.PDF_FIT, Optional ByRef leftX As Variant = Null, Optional ByRef rightX As Variant = Null, Optional ByRef topY As Variant = Null, Optional ByRef bottomY As Variant = Null, Optional ByRef zoom As Variant = Null, Optional ByRef extra As Dictionary = Nothing) As pdfValue
```

**Parameters:**
- `page` (Long): Target page number.
- `fit` (PDF_FIT, Optional): Fit mode. Default is PDF_FIT.
- `leftX` (Variant, Optional): Left X coordinate for applicable fit modes.
- `rightX` (Variant, Optional): Right X coordinate for applicable fit modes.
- `topY` (Variant, Optional): Top Y coordinate for applicable fit modes.
- `bottomY` (Variant, Optional): Bottom Y coordinate for applicable fit modes.
- `zoom` (Variant, Optional): Zoom factor for PDF_XYZ fit mode.
- `extra` (Dictionary, Optional): Additional destination properties.

**Returns:**
- `pdfValue`: New destination object.

**Usage:**
```vba
' Fit entire page
Dim dest1 As pdfValue
Set dest1 = doc.NewDestination(1, PDF_FIT.PDF_FIT)

' Fit to specific coordinates with zoom
Dim dest2 As pdfValue
Set dest2 = doc.NewDestination(2, PDF_FIT.PDF_XYZ, 100, , 200, , 1.5)
```

#### parseDestination Subroutine
Parses a destination object into its components.

```vba
Public Sub parseDestination(ByRef dest As pdfValue, ByRef page As Long, ByRef fit As PDF_FIT, ByRef leftX As Variant, ByRef rightX As Variant, ByRef topY As Variant, ByRef bottomY As Variant, ByRef zoom As Variant, ByRef extra As Dictionary)
```

**Parameters:**
- `dest` (pdfValue): Destination object to parse.
- `page` (Long): Returns target page number.
- `fit` (PDF_FIT): Returns fit mode.
- `leftX` (Variant): Returns left X coordinate.
- `rightX` (Variant): Returns right X coordinate.
- `topY` (Variant): Returns top Y coordinate.
- `bottomY` (Variant): Returns bottom Y coordinate.
- `zoom` (Variant): Returns zoom factor.
- `extra` (Dictionary): Returns additional properties.

### Document Outlines (Bookmarks)

#### NewOutlines Function
Creates a new document outline tree.

```vba
Public Function NewOutlines(ByRef parent As pdfValue, Optional ByRef defaults As Dictionary = Nothing) As pdfValue
```

**Parameters:**
- `parent` (pdfValue): Parent object reference.
- `defaults` (Dictionary, Optional): Additional default values.

**Returns:**
- `pdfValue`: New outline tree object.

#### AddOutlines Subroutine
Adds or initializes the document outline tree.

```vba
Public Sub AddOutlines(Optional ByRef anOutlineItem As pdfValue = Nothing)
```

**Parameters:**
- `anOutlineItem` (pdfValue, Optional): Outline item to set. If Nothing, creates default outline tree.

#### NewOutlineItem Function
Creates a new outline item (bookmark).

```vba
Public Function NewOutlineItem(ByRef parent As pdfValue, Optional ByRef defaults As Dictionary) As pdfValue
```

**Parameters:**
- `parent` (pdfValue): Parent outline object.
- `defaults` (Dictionary, Optional): Additional default values including /Title, /Dest, etc.

**Returns:**
- `pdfValue`: New outline item object.

#### AddOutlineItem Subroutine
Adds an outline item to the document outline tree.

```vba
Public Sub AddOutlineItem(ByRef parent As pdfValue, Optional ByRef anOutlineItem As pdfValue = Nothing)
```

**Parameters:**
- `parent` (pdfValue): Parent outline object. If Nothing, uses root outline.
- `anOutlineItem` (pdfValue, Optional): Outline item to add.

**Usage:**
```vba
' Initialize outlines
doc.AddOutlines

' Create bookmark with title and destination
Dim defaults As New Dictionary
defaults("/Title") = "Chapter 1"
Dim dest As pdfValue
Set dest = doc.NewDestination(1, PDF_FIT.PDF_FIT)
defaults("/Dest") = dest

Dim bookmark As pdfValue
Set bookmark = doc.NewOutlineItem(doc.Outlines, defaults)
doc.AddOutlineItem doc.Outlines, bookmark
```

### Object Management

#### getObject Function
Retrieves a PDF object by ID and generation.

```vba
Public Function getObject(ByVal id As Long, ByVal generation As Long, Optional ByVal cacheObject As Boolean = True) As pdfValue
```

**Parameters:**
- `id` (Long): Object ID number.
- `generation` (Long): Object generation number.
- `cacheObject` (Boolean, Optional): Whether to cache the object. Default is True.

**Returns:**
- `pdfValue`: Retrieved object or null object if not found.

**Usage:**
```vba
Dim obj As pdfValue
Set obj = doc.getObject(5, 0)
If obj.valueType <> PDF_ValueType.PDF_Null Then
    Debug.Print "Found object 5"
End If
```

#### getCachedObject Function
Retrieves a PDF object from cache only.

```vba
Public Function getCachedObject(ByVal id As Long, ByVal generation As Long) As pdfValue
```

**Parameters:**
- `id` (Long): Object ID number.
- `generation` (Long): Object generation number.

**Returns:**
- `pdfValue`: Cached object or null object if not in cache.

#### renumberIds Function
Renumbers all object IDs starting from a base ID.

```vba
Public Function renumberIds(ByVal baseId As Long, Optional ByRef root As pdfValue = Nothing, Optional ByRef visited As Dictionary = Nothing) As Long
```

**Parameters:**
- `baseId` (Long): Starting ID number for renumbering.
- `root` (pdfValue, Optional): Root object to start from. If Nothing, renumbers entire document.
- `visited` (Dictionary, Optional): Internal tracking dictionary.

**Returns:**
- `Long`: Next available ID after renumbering.

**Usage:**
```vba
' Renumber all objects starting from ID 1
Dim nextId As Long
nextId = doc.renumberIds(1)
Debug.Print "Next available ID: " & nextId
```

## Module-Level API (Neither Public nor Private)

### PDF Content Parsing

#### GetValueType Function
Determines the PDF value type at a specific offset in byte array.

```vba
Function GetValueType(ByRef bytes() As Byte, ByVal offset As Long) As PDF_ValueType
```

**Parameters:**
- `bytes()` (Byte Array): PDF content bytes.
- `offset` (Long): Position in byte array.

**Returns:**
- `PDF_ValueType`: Type of value found at offset.

#### GetValue Function
Parses and returns a PDF value from byte array.

```vba
Function GetValue(ByRef bytes() As Byte, ByRef offset As Long, Optional ByRef Meta As Dictionary = Nothing) As pdfValue
```

**Parameters:**
- `bytes()` (Byte Array): PDF content bytes.
- `offset` (Long): Position in byte array (updated after parsing).
- `Meta` (Dictionary, Optional): Metadata for stream objects.

**Returns:**
- `pdfValue`: Parsed PDF value object.

#### loadObject Function
Loads a PDF object from raw content using cross-reference table.

```vba
Function loadObject(ByVal Index As Long, Optional forceReload As Boolean = False) As pdfValue
```

**Parameters:**
- `Index` (Long): Object ID to load.
- `forceReload` (Boolean, Optional): Force reload even if cached. Default is False.

**Returns:**
- `pdfValue`: Loaded object or null if not found.

### Document Structure Functions

#### GetTrailer Function
Extracts the PDF trailer from content.

```vba
Function GetTrailer(ByRef content() As Byte) As pdfValue
```

**Parameters:**
- `content()` (Byte Array): PDF file content.

**Returns:**
- `pdfValue`: Trailer object or null if not found.

#### GetRootObject Function
Retrieves the document catalog (root object).

```vba
Function GetRootObject(ByRef content() As Byte, ByRef trailer As pdfValue, ByRef xrefTable As Dictionary) As pdfValue
```

**Parameters:**
- `content()` (Byte Array): PDF file content.
- `trailer` (pdfValue): Document trailer.
- `xrefTable` (Dictionary): Cross-reference table.

**Returns:**
- `pdfValue`: Root catalog object.

#### GetXrefTable Function
Extracts the cross-reference table from PDF content.

```vba
Function GetXrefTable(ByRef content() As Byte, ByRef trailer As pdfValue) As Dictionary
```

**Parameters:**
- `content()` (Byte Array): PDF file content.
- `trailer` (pdfValue): Document trailer.

**Returns:**
- `Dictionary`: Cross-reference table with xrefEntry objects.

#### ParseXrefTable Function
Parses cross-reference table data.

```vba
Function ParseXrefTable(ByRef content() As Byte, ByRef offset As Long, ByRef trailer As pdfValue, Optional ByRef xrefTable As Dictionary = Nothing) As Dictionary
```

**Parameters:**
- `content()` (Byte Array): PDF file content.
- `offset` (Long): Starting position for parsing.
- `trailer` (pdfValue): Document trailer.
- `xrefTable` (Dictionary, Optional): Existing table to append to.

**Returns:**
- `Dictionary`: Parsed cross-reference table.

### Object Tree Processing

#### GetObjectsInTree Subroutine
Recursively loads all objects referenced from a root object.

```vba
Sub GetObjectsInTree(ByRef root As pdfValue, ByRef content() As Byte, ByRef xrefTable As Dictionary, ByRef objects As Dictionary)
```

**Parameters:**
- `root` (pdfValue): Starting object for traversal.
- `content()` (Byte Array): PDF file content.
- `xrefTable` (Dictionary): Cross-reference table.
- `objects` (Dictionary): Dictionary to populate with found objects.

### PDF Writing Functions

#### SavePdfHeader Function
Writes PDF header to file and returns file handle.

```vba
Function SavePdfHeader(ByRef pdfFilename As String, ByRef offset As Long, Optional ByVal Header As String) As Integer
```

**Parameters:**
- `pdfFilename` (String): Output filename.
- `offset` (Long): Returns current file position.
- `Header` (String, Optional): PDF header string.

**Returns:**
- `Integer`: File handle for continued writing.

#### SavePdfObject Subroutine
Saves a single PDF object to file.

```vba
Sub SavePdfObject(ByRef outputFileNum As Integer, ByRef obj As pdfValue, ByRef offset As Long, Optional ByVal baseId As Long = 0, Optional ByVal prettyPrint As Boolean = True)
```

**Parameters:**
- `outputFileNum` (Integer): File handle.
- `obj` (pdfValue): Object to save.
- `offset` (Long): Current file position (updated).
- `baseId` (Long, Optional): Base ID for renumbering.
- `prettyPrint` (Boolean, Optional): Format output nicely.

#### SavePdfObjects Subroutine
Saves all objects in a dictionary to file.

```vba
Sub SavePdfObjects(ByRef outputFileNum As Integer, ByRef pdfObjs As Dictionary, ByRef offset As Long, Optional ByVal baseId As Long = 0)
```

**Parameters:**
- `outputFileNum` (Integer): File handle.
- `pdfObjs` (Dictionary): Objects to save.
- `offset` (Long): Current file position (updated).
- `baseId` (Long, Optional): Base ID for renumbering.

#### SaveOutlinePdfObjects Subroutine
Saves outline hierarchy to file.

```vba
Sub SaveOutlinePdfObjects(ByRef outputFileNum As Integer, ByRef outlineParentObj As pdfValue, ByRef offset As Long, Optional ByVal baseId As Long = 0, Optional ByVal prettyPrint As Boolean = True)
```

**Parameters:**
- `outputFileNum` (Integer): File handle.
- `outlineParentObj` (pdfValue): Root outline object.
- `offset` (Long): Current file position (updated).
- `baseId` (Long, Optional): Base ID for renumbering.
- `prettyPrint` (Boolean, Optional): Format output nicely.

#### SavePdfTrailer Subroutine
Saves trailer, cross-reference table, and closes file.

```vba
Sub SavePdfTrailer(ByRef outputFileNum As Integer, ByRef offset As Long, Optional ByVal prettyPrint As Boolean = True)
```

**Parameters:**
- `outputFileNum` (Integer): File handle to close.
- `offset` (Long): Current file position.
- `prettyPrint` (Boolean, Optional): Format output nicely.

### Utility Functions

#### NewTrailer Function
Creates a default trailer object.

```vba
Function NewTrailer() As pdfValue
```

**Returns:**
- `pdfValue`: New trailer object.

#### NewXrefTable Function
Creates a default cross-reference table.

```vba
Function NewXrefTable() As Dictionary
```

**Returns:**
- `Dictionary`: New cross-reference table with required entry 0.

#### AddUpdateXref Subroutine
Updates cross-reference table with object information.

```vba
Sub AddUpdateXref(ByRef obj As pdfValue, ByVal offset As Long, Optional ByVal baseId As Long = 0)
```

**Parameters:**
- `obj` (pdfValue): Object to add/update.
- `offset` (Long): File offset of object.
- `baseId` (Long, Optional): Base ID for renumbering.

## Private Internal Helper Routines

### Name Processing

#### ProcessName Function
Processes PDF name values, handling hex encoding and UTF-8.

```vba
Public Function ProcessName(ByRef name() As Byte) As String
```

**Parameters:**
- `name()` (Byte Array): Raw name bytes from PDF.

**Returns:**
- `String`: Processed name string with proper encoding.

#### EscapeName Function
Escapes special characters in names for PDF output.

```vba
Public Function EscapeName(ByRef name As String, Optional ByVal addUtf8BOM As Boolean = False) As String
```

**Parameters:**
- `name` (String): Name to escape.
- `addUtf8BOM` (Boolean, Optional): Whether to add UTF-8 BOM.

**Returns:**
- `String`: Escaped name suitable for PDF output.

### Object Management Helpers

#### updateTrailerReference Subroutine
Updates object references in the trailer dictionary.

```vba
Private Sub updateTrailerReference(ByRef key As String, ByRef valueObj As pdfValue)
```

**Parameters:**
- `key` (String): Dictionary key to update.
- `valueObj` (pdfValue): Object to reference.

#### objFromDocCatalog Function
Retrieves objects from the document catalog with caching.

```vba
Private Function objFromDocCatalog(ByRef m_obj As pdfValue, ByRef keyName As String, Optional ByVal isOptional As Boolean = True) As pdfValue
```

**Parameters:**
- `m_obj` (pdfValue): Cached object reference.
- `keyName` (String): Key name in catalog.
- `isOptional` (Boolean, Optional): Whether object is optional.

**Returns:**
- `pdfValue`: Retrieved object from catalog.

#### NewTopLevelDictionary Function
Creates top-level dictionary objects with proper initialization.

```vba
Private Function NewTopLevelDictionary(ByRef typeName As String, ByRef parent As pdfValue, Optional ByRef defaults As Dictionary = Nothing) As pdfValue
```

**Parameters:**
- `typeName` (String): PDF type name (e.g., "/Pages").
- `parent` (pdfValue): Parent object reference.
- `defaults` (Dictionary, Optional): Additional default values.

**Returns:**
- `pdfValue`: New dictionary object.

#### GetEndOfChain Function
Follows linked list to find the end object.

```vba
Private Function GetEndOfChain(ByRef item As pdfValue, ByRef linkName As String) As pdfValue
```

**Parameters:**
- `item` (pdfValue): Starting object in chain.
- `linkName` (String): Link property name (e.g., "/Next").

**Returns:**
- `pdfValue`: Reference to last object in chain.

### Low-Level Parsing Helpers

#### getInt Function
Reads integer values from byte array.

```vba
Private Function getInt(ByRef content() As Byte, ByRef offset As Long, ByVal byteCount As Long, ByVal defaultValue As Long) As Long
```

**Parameters:**
- `content()` (Byte Array): Source bytes.
- `offset` (Long): Position to read from (updated).
- `byteCount` (Long): Number of bytes to read.
- `defaultValue` (Long): Default if byteCount is 0.

**Returns:**
- `Long`: Integer value read from bytes.

### Class Lifecycle

#### Class_Initialize Subroutine
Initializes new pdfDocument instance with default objects.

```vba
Private Sub Class_Initialize()
```

#### Class_Terminate Subroutine
Cleans up resources when object is destroyed.

```vba
Private Sub Class_Terminate()
```

## Usage Examples

### Creating a New PDF Document

```vba
Sub CreateNewPDF()
    Dim doc As pdfDocument
    Set doc = pdfDocument.pdfDocument()

    ' Set document properties
    doc.version = PDF_VERSIONS.PDF_1_7
    doc.Title = "My New Document"

    ' Add document info
    doc.AddInfo
    With doc.Info.asDictionary()
        .Item("/Author") = "John Doe"
        .Item("/Subject") = "Test Document"
        .Item("/Creator") = "VBA PDF Library"
    End With

    ' Initialize pages
    doc.AddPages

    ' Add a page
    Dim newPage As pdfValue
    Set newPage = doc.NewPage(doc.Pages)
    doc.AddPages newPage

    ' Save document
    If doc.savePdfAs("C:\output\new_document.pdf") Then
        Debug.Print "Document created successfully"
    End If
End Sub
```

### Loading and Modifying an Existing PDF

```vba
Sub ModifyExistingPDF()
    Dim doc As pdfDocument
    Set doc = pdfDocument.pdfDocument("C:\input\document.pdf")

    If Not doc Is Nothing Then
        Debug.Print "Loaded PDF with " & doc.pageCount & " pages"

        ' Modify title
        doc.Title = "Modified: " & doc.Title

        ' Add bookmark
        doc.AddOutlines
        Dim defaults As New Dictionary
        defaults("/Title") = "First Page"
        Dim dest As pdfValue
        Set dest = doc.NewDestination(1, PDF_FIT.PDF_FIT)
        defaults("/Dest") = dest

        Dim bookmark As pdfValue
        Set bookmark = doc.NewOutlineItem(doc.Outlines, defaults)
        doc.AddOutlineItem doc.Outlines, bookmark

        ' Save modified document
        doc.savePdfAs "C:\output\modified_document.pdf"
    End If
End Sub
```

### Working with Named Destinations

```vba
Sub AddNamedDestinations()
    Dim doc As pdfDocument
    Set doc = pdfDocument.pdfDocument("document.pdf")

    If Not doc Is Nothing Then
        ' Create destinations for each page
        Dim i As Long
        For i = 1 To doc.pageCount
            Dim destName As pdfValue
            Set destName = pdfValue.NewNameValue("/Page" & i)

            Dim dest As pdfValue
            Set dest = doc.NewDestination(i, PDF_FIT.PDF_FITH, , , 800)

            doc.AddNamedDestinations destName, dest
        Next i

        doc.savePdf
    End If
End Sub
```
