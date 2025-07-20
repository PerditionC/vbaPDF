# vbaPDF API Documentation

A VBA library for reading, writing, and modifying Portable Document Format (PDF) files.

**Note:** This library supports smaller PDF documents that fully fit in memory, as the full document is read at once rather than streaming or loading portions on demand.

## Table of Contents

- [Getting Started](#getting-started)
- [Core Classes](#core-classes)
- [Enumerations](#enumerations)
- [Loading and Saving](#loading-and-saving)
- [Properties](#properties)
- [Document Structure](#document-structure)
- [Page Management](#page-management)
- [Destinations and Navigation](#destinations-and-navigation)
- [Document Outlines (Bookmarks)](#document-outlines-bookmarks)
- [Text Extraction](#text-extraction)
- [Low Level API](#low-level-api)
- [Examples](#examples)

## Getting Started

### Basic Usage

```vba
Dim pdf As pdfDocument
Set pdf = New pdfDocument

' Open, read, and then parse an existing PDF
pdf.OpenPdf "C:\path\to\document.pdf"

' Access document properties
Debug.Print "Title: " & pdf.Title
Debug.Print "Page Count: " & pdf.pageCount

' Save changes
pdf.SavePdfAs "C:\path\to\output.pdf"
```

## Core Classes

### pdfValue

Class which represents any value stored in a pdf file.  Only pdf object values have a valid id and generation.

### pdfDocument

The main class for working with PDF documents. Represents the complete PDF document structure including header, objects, catalog, and trailer.

#### pdfDocument Function
Creates a new pdfDocument instance with optional file loading.

```vba
Public Function pdfDocument(Optional ByVal pdfFilename As String = vbNullString) As pdfDocument
```

**Parameters:**
- `pdfFilename` (String, Optional): Path to PDF file to load. If empty, creates blank document.

**Returns:** 
- `pdfDocument`: New document instance, or Nothing if loading fails.

**Example:**
```vba
' Create new blank document
Dim pdf As pdfDocument
Set pdf = pdfDocument.pdfDocument()

' Open (Load and Parse) existing PDF
Set pdf = pdfDocument.pdfDocument("C:\path\to\file.pdf")
If pdf Is Nothing Then
    Debug.Print "Failed to load PDF"
End If
```

## Enumerations

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
    PDF_1_7 = 17
    PDF_2_0 = 20
    PDF_Default = PDF_1_7
End Enum
```

### PDF_FIT

Defines page fit modes for destinations and viewing.

```vba
Public Enum PDF_FIT
    PDF_XYZ = -1    ' Display with specified coordinates and zoom
    PDF_FIT = 0     ' Fit entire page in window
    PDF_FITH        ' Fit page width, specified top coordinate
    PDF_FITV        ' Fit page height, specified left coordinate
    PDF_FITR        ' Fit specified rectangle
    PDF_FITB        ' Fit bounding box of page contents
    PDF_FITBH       ' Fit width of bounding box, specified top
    PDF_FITBV       ' Fit height of bounding box, specified left
End Enum
```

## Loading and Saving

### OpenPdf Function
Loads and Parses specified PDF file

```vba
Public Function OpenPdf(ByVal pdfFilename As String) As Boolean
```

**Parameters:**
- `pdfFilename` (String): Path to PDF file to load.  An error if file does not exist.

**Returns:**
- `Boolean`: True on success, False on error.

**Usage:**
```vba
Dim pdf As pdfDocument: Set pdf = New pdfDocument
If pdf.OpenPdf("C:\Documents\sample.pdf") Then
    Debug.Print "PDF loaded and parsed successfully"
    Debug.Print "Page count: " & doc.pageCount
End If
```

### LoadPdf Function
Loads a PDF document from file.  Reads in file and cross reference table, does **not** load any pdf objects.

```vba
Public Function loadPdf(ByVal pdfFilename As String) As Boolean
```

**Parameters:**
- `pdfFilename` (String): Full path to PDF file to load.  An error is file does not exist.

**Returns:**
- `Boolean`: True on success, False on error.

**Example:**
```vba
Dim pdf As pdfDocument: Set pdf = New pdfDocument
If pdf.LoadPdf("C:\Documents\sample.pdf") Then
    Debug.Print "PDF loaded successfully"
Else
    Debug.Print "Failed to load PDF"
End If
```

### ParsePdf Function
Parses (reads pdf objects from) the loaded PDF document and populates object cache.

```vba
Public Function ParsePdf() As Boolean
```

**Returns:**
- `Boolean`: True on success, False on error.

**Usage:**
```vba
Dim pdf As pdfDocument: Set pdf = New pdfDocument
If pdf.loadPdf("document.pdf") Then
    If pdf.ParsePdf() Then
        Debug.Print "PDF parsed successfully"
        Debug.Print "Page count: " & doc.pageCount
    End If
End If
```

### `SavePdf Function`
Saves the PDF document, replacing original.

```vba
Public Function SavePdf() As Boolean
```

**Returns:**
- `Boolean`: True on success, False on error.

**Example:**
```vba
pdf.Title = "Modified Document"
If pdf.SavePdf() Then
    Debug.Print "Document saved"
End If
```

### SavePdfAs Function
Saves the PDF document to specified file.

```vba
Function SavePdfAs(pdfFilename As String) As Boolean
```

**Parameters:**
- `pdfFilename` (String): Full path where the PDF should be saved

**Returns:**
- `Boolean`: True on success, False on error.

**Example:**
```vba
If pdf.SavePdfAs("C:\Documents\modified.pdf") Then
    Debug.Print "Document saved to new file: " & pdf.filename
Else
    Debug.Print "Error Saving PDF"
End If
```

## Properties

### Version Property
Gets or sets the PDF version. Changes to version also update the header.

```vba
Public Property Let Version(ByVal pdfVersion As PDF_VERSIONS)
Public Property Get Version() As PDF_VERSIONS
```

**Parameters:**
- `pdfVersion` (PDF_VERSIONS): Version to set.

**Returns:**
- `PDF_VERSIONS`: Current PDF version.

**Example:**
```vba
pdf.version = PDF_VERSIONS.PDF_1_7
Debug.Print "PDF Version: " & pdf.version
```

### Header Property
Gets or sets the PDF header string. Changes to header also update the version.

```vba
Public Property Let Header(ByVal pdfHeader As String)
Public Property Get Header() As String
```

**Parameters:**
- `pdfHeader` (String): Header string to set (e.g., "%PDF-1.7").

**Returns:**
- `String`: Current PDF header.

**Example:**
```vba
pdf.Header = "%PDF-1.7" & vbNewLine
Debug.Print pdf.Header
```

### Object ID Management

#### nextObjId
Gets or sets the next available object id for creating new PDF objects.  See also pdfValue.id

```vba
Public Property Let nextObjId(ByVal nextId As Long)
Public Property Get nextObjId() As Long
```

**Parameters:**
- `nextId` (Long): Next id to use.

**Returns:**
- `Long`: Next available object id (increments on each call).

**Example:**
```vba
Dim newId As Long
newId = pdf.nextObjId
Debug.Print "Next Object id: " & newId
```

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

**Example:**
```vba
Dim obj As pdfValue
Set obj = pdf.getObject(5, 0)
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
nextId = pdf.renumberIds(1)
Debug.Print "Next available ID: " & nextId
```

### Document Information

#### Info Property
Gets or sets the document information dictionary (/Info). Updates /Info in trailer.

```vba
Public Property Set Info(ByRef Info As pdfValue)
Public Property Get Info() As pdfValue
```

**Parameters:**
- `Info` (pdfValue): Information dictionary to set.

**Returns:**
- `pdfValue`: Current document information dictionary.

**Example:**
```vba
' Get existing info
Dim docInfo As pdfValue
Set docInfo = pdf.Info

' Add new info
pdf.AddInfo
Set pdf.Info.asDictionary.item("/Author") = pdfValue.NewValue("John Doe")

' Display values from info
If pdf.Info.hasKey("/Author") Then
    Debug.Print "Author: " & pdf.Info.GetValue("/Author")
End If
```

#### Meta Property
Gets or sets the document metadata object (/Metadata). Updates /Metadata in document catalog.

```vba
Public Property Set Meta(ByRef Meta As pdfValue)
Public Property Get Meta() As pdfValue
```

**Parameters:**
- `Meta` (pdfValue): Metadata stream to set.

**Returns:**
- `pdfValue`: Current document metadata stream.

### Title Property
Gets or sets the document title. Returns filename if no title is set in /Info.

```vba
Public Property Let Title(pdfTitle As String)
Public Property Get Title() As String
```

**Parameters:**
- `pdfTitle` (String): Title to set. Empty string removes title.

**Returns:**
- `String`: Current document title or filename if no title set.

**Example:**
```vba
pdf.Title = "My Document Title"
Debug.Print "Document Title: " & pdf.Title
```

### Document Structure

#### Pages Property
Gets or sets the document pages tree root object.

```vba
Public Property Set Pages(ByRef Pages As pdfValue)
Public Property Get Pages() As pdfValue
```

**Parameters:**
- `Pages` (pdfValue): Pages tree to set.

**Returns:**
- `pdfValue`: Current pages tree object.

#### PageCount Property (Read-Only)
Returns the total number of pages in the document.

```vba
Public Property Get PageCount() As Long
```

**Returns:**
- `Long`: Number of pages in document.

**Example:**
```vba
Debug.Print "Document has " & pdf.PageCount & " pages"
```

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
Gets or sets the document outline (bookmarks) tree.

```vba
Public Property Set Outlines(ByRef Outlines As pdfValue)
Public Property Get Outlines() As pdfValue
```

**Parameters:**
- `Outlines` (pdfValue): Outline tree to set.

**Returns:**
- `pdfValue`: Current outline tree object.

## Document Structure

### Document Creation and Modification

#### AddInfo Subroutine
Adds and initializes the document information dictionary, /Info in the trailer.

```vba
Public Sub AddInfo(Optional defaults As Variant)
```

**Parameters:**
- `defaults` (Variant, Optional): Additional default values to include.

**Example:**
```vba
pdf.AddInfo ' Automatically adds /Producer = "vbaPDF"

Set pdf.Info.asDictionary()("/Author") = pdfValue.NewValue("John Doe")
Set pdf.Info.asDictionary()("/Subject") = pdfValue.NewValue("Test Document")
```

#### NewDocumentCatalog Subroutine
Creates and initializes the top-level /Root document catalog object.

```vba
Public Sub NewDocumentCatalog()
```

**Example:**
```vba
Dim pdf As pdfDocument: Set pdf = pdfDocument.pdfDocument()
pdf.NewDocumentCatalog()
' Creates catalog with default values:
' /Type = /Catalog
' /Version = current PDF version
' /PageLayout = /OneColumn
' /PageMode = /UseNone
' /Lang = "en"
```

## Page Management

#### NewPages Function
Creates a new /Pages tree object with default initialization.

```vba
Public Function NewPages(ByRef parent As pdfValue, Optional ByRef defaults As Dictionary = Nothing) As pdfValue
```

**Parameters:**
- `parent` (pdfValue): Parent pages object reference (Nothing for root)
- `defaults` (Dictionary, Optional): Additional default values.

**Returns:**
- `pdfValue`: New /Pages tree object.

**Example:**
```vba
Dim rootPages As pdfValue
Set rootPages = pdf.NewPages(Nothing)
```

#### NewPage Function
Creates a new /Page object.

```vba
Public Function NewPage(ByRef parent As pdfValue, Optional ByRef defaults As Dictionary = Nothing) As pdfValue
```

**Parameters:**
- `parent` (pdfValue): Parent pages tree reference.
- `defaults` (Dictionary, Optional): Additional default values.

**Returns:**
- `pdfValue`: New pdfValue representing the /Page object

**Example:**
```vba
Dim newPage As pdfValue
Set newPage = pdf.NewPage(pdf.Pages)
```

```vba
Dim pagesTree As pdfValue
Set pagesTree = doc.NewPages(Nothing)
Dim newPage As pdfValue
Set newPage = doc.NewPage(pagesTree)
```

#### AddPages Subroutine
Adds a /Page or /Pages object to the document. If pages is Nothing, initializes the top-level /Pages.

```vba
Public Sub AddPages(Optional ByRef thePages As pdfValue = Nothing)
```

**Parameters:**
- `thePages` (pdfValue, Optional): The page or pages object to add. If Nothing, initializes top-level pages tree.

**Example:**
```vba
' Initialize root pages tree
pdf.AddPages

' Add a new page
Dim page As pdfValue
Set page = pdf.NewPage(pdf.Pages)
pdf.AddPages page
```

## Destinations and Navigation

### Creating Destinations

#### NewDestination Function
Creates a new destination object for navigation.

```vba
Public Function NewDestination(ByRef page As Long, Optional ByRef fit As PDF_FIT = PDF_FIT.PDF_FIT, Optional ByRef leftX As Variant = Null, Optional ByRef rightX As Variant = Null, Optional ByRef topY As Variant = Null, Optional ByRef bottomY As Variant = Null, Optional ByRef zoom As Variant = Null, Optional ByRef extra As Dictionary = Nothing) As pdfValue
```

**Parameters:**
- `page` (Long): Target page number.
- `fit` (PDF_FIT, Optional): Fit mode. Default is PDF_FIT (see PDF_FIT enumeration).
- `leftX`, `rightX`, `topY`, `bottomY`: Position coordinates (depending on fit mode)
- `zoom` (Variant, Optional): Zoom factor (for PDF_XYZ fit mode).
- `extra` (Dictionary, Optional): Additional destination dictionary properties.

**Returns:**
- `pdfValue`: object representing the destination.

**Example:**
```vba
' Create destination to fit entire first page
Dim dest1 As pdfValue
Set dest1 = pdf.NewDestination(0, PDF_FIT.PDF_FIT)

' Create destination with fit to specific coordinates and zoom on page 3
Dim dest2 As pdfValue
Set dest2 = pdf.NewDestination(2, PDF_FIT.PDF_XYZ, 100, , 200, , 1.5)
```

#### parseDestination Subroutine
Parses a destination object into its component parts.

```vba
Public Sub parseDestination(ByRef dest As pdfValue, ByRef page As Long, ByRef fit As PDF_FIT, ByRef leftX As Variant, ByRef rightX As Variant, ByRef topY As Variant, ByRef bottomY As Variant, ByRef zoom As Variant, ByRef extra As Dictionary)
```

**Parameters:**
- `dest` (pdfValue): Destination object to parse.
- `page` (Long, Output): Returns target page number.
- `fit` (PDF_FIT, Output): Returns fit mode.
- `leftX`, `rightX`, `topY`, `bottomY`: (Variant, Output) Position coordinates
- `zoom` (Variant, Output): Returns zoom factor.
- `extra` (Dictionary, Output): Returns dictionary of additional property values.

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
- `destName` (pdfValue of type PDF_Name): Name of the destination.
- `theDest` (pdfValue): Destination array or dictionary object.

**Example:**
```vba
' Create named destination
Dim destName As pdfValue
Set destName = pdfValue.NewNameValue("/MyDest")

Dim dest As pdfValue
Set dest = pdf.NewDestination(1, PDF_FIT.PDF_FIT)

pdf.AddNamedDestinations destName, dest
```

```vba
' Create named destination
Dim destName As pdfValue
Set destName = pdfValue.NewNameValue("/Chapter1")

Dim dest As pdfValue
Set dest = pdf.NewDestination(5, PDF_FIT.PDF_FITH, , , 100)

pdf.AddNamedDestinations destName, dest
```

## Document Outlines (Bookmarks)

#### NewOutlines Function
Creates a new document outline (bookmarks) tree.

```vba
Public Function NewOutlines(ByRef parent As pdfValue, Optional ByRef defaults As Dictionary = Nothing) As pdfValue
```

**Parameters:**
- `parent` (pdfValue): Parent object reference.
- `defaults` (Dictionary, Optional): Additional default values.

**Returns:**
- `pdfValue`: New outline tree object.

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

#### AddOutlines Subroutine
Adds or initializes the document outline tree.

```vba
Public Sub AddOutlines(Optional ByRef anOutlineItem As pdfValue = Nothing)
```

**Parameters:**
- `anOutlineItem` (pdfValue, Optional): Outline item to set. If Nothing, creates default outline tree.

#### AddOutlineItem Subroutine
Adds an outline item to a parent outline tree.

```vba
Public Sub AddOutlineItem(ByRef parent As pdfValue, Optional ByRef anOutlineItem As pdfValue = Nothing)
```

**Parameters:**
- `parent` (pdfValue): Parent outline object. If Nothing, uses document root outline.
- `anOutlineItem` (pdfValue, Optional): Outline item to add.

**Example:**
```vba
' Initialize outlines
pdf.AddOutlines

' Create bookmark (outline) with title and destination
Dim defaults As Dictionary
Set defaults = New Dictionary
defaults.Add "/Title", pdfValue.NewValue("Chapter 1")

Dim dest As pdfValue
Set dest = pdf.NewDestination(1, PDF_FIT.PDF_FIT)
Set defaults("/Dest") = dest

Dim bookmark As pdfValue
Set bookmark = doc.NewOutlineItem(pdf.Outlines, defaults)
pdf.AddOutlineItem doc.Outlines, bookmark
```

## Text Extraction

### Public Text Extraction API

#### `Function GetPageText(PageIndex As Long) As String`
Extracts all visible text from a single page.

#### `Function GetDocumentText() As String`
Extracts all text from the entire document.

#### `Function GetPageTextWithLayout(PageIndex As Long, Optional AsHtml As Boolean = False) As Variant`
Extracts text with layout and formatting information.

#### `Sub SetFontUnicodeMap(FontName As String, Map As Dictionary)`
Sets font encoding mappings for custom or subset fonts.

### Internal Text API

#### `Function ParseContentStream(PageIndex As Long) As Collection`
Parses raw content stream operations.

#### `Function MapFontCodeToUnicode(FontName As String, Code As Variant) As String`
Maps font codes to Unicode characters.

## Low Level API
(Internal / Module)

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

## Examples

### Creating a New PDF Document

```vba
Sub CreateNewPDF()
    Dim pdf As pdfDocument
    Set pdf = pdfDocument.pdfDocument()

    ' Set version and create document structure
    pdf.version = PDF_VERSIONS.PDF_1_7
    pdf.NewDocumentCatalog
    pdf.AddInfo

    ' Set document properties
    pdf.Title = "My New Document"
    With pdf.Info.asDictionary()
        .Item("/Author") = pdfValue.NewValue("John Doe")
        .Item("/Subject") = pdfValue.NewValue("Test Document")
        .Item("/Creator") = pdfValue.NewValue("VBA PDF Library")
    End With

    ' Initialize pages
    pdf.AddPages

    ' Add a page
    Dim page As pdfValue
    Set page = pdf.NewPage(pdf.Pages)
    pdf.AddPages page

    ' Save the document
    If pdf.SavePdfAs("C:\temp\new_document.pdf") Then
        Debug.Print "Document created successfully"
    End If
End Sub
```

### Reading PDF Information

```vba
Sub ReadPDFInfo()
    Dim pdf As pdfDocument
    Set pdf = pdfDocument.pdfDocument()

    pdf.OpenPdf "C:\Documents\sample.pdf"

    ' Display document information
    Debug.Print "Title: " & pdf.Title
    Debug.Print "Page Count: " & pdf.PageCount
    Debug.Print "PDF Version: " & pdf.Version

    ' Access detailed info if available
    If pdf.Info.hasKey("/Author") Then
        Debug.Print "Author: " & pdf.Info.GetValue("/Author").value
    End If

    If pdf.Info.hasKey("/CreationDate") Then
        Debug.Print "Created: " & pdf.Info.asDictionary.Item("/CreationDate").value
    End If
End Sub
```

### Loading and Modifying an Existing PDF

```vba
Sub ModifyExistingPDF()
    Dim pdf As pdfDocument
    Set pdf = pdfDocument.pdfDocument("C:\temp\document.pdf")

    If Not pdf Is Nothing Then
        Debug.Print "Loaded PDF with " & pdf.PageCount & " pages"

        ' Modify title
        pdf.Title = "Modified: " & pdf.Title

        ' Add bookmark
        pdf.AddOutlines
        Dim defaults As New Dictionary
        Set defaults("/Title") = pdfValue.NewValue("First Page")
        Dim dest As pdfValue
        Set dest = pdf.NewDestination(0, PDF_FIT.PDF_FIT)
        Set defaults("/Dest") = dest

        Dim bookmark As pdfValue
        Set bookmark = pdf.NewOutlineItem(pdf.Outlines, defaults)
        pdf.AddOutlineItem pdf.Outlines, bookmark

        ' Save modified document
        pdf.savePdfAs "C:\temp\modified_document.pdf"
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

```vba
Sub CreateNamedDestinations()
    Dim pdf As pdfDocument
    Set pdf = New pdfDocument

    pdf.OpenPdf "C:\Documents\sample.pdf"

    ' Create a named destination to page 3, fit width
    Dim destName As pdfValue
    Set destName = pdfValue.NewNameValue("/Chapter2")

    Dim dest As pdfValue
    Set dest = pdf.NewDestination(2, PDF_FIT.PDF_FITH, , , 100)

    pdf.AddNamedDestinations destName, dest

    pdf.SavePdfAs "C:\Documents\sample_with_destinations.pdf"
End Sub
```

### Creating Document Outlines

```vba
Sub CreateOutlines()
    Dim pdf As pdfDocument
    Set pdf = New pdfDocument

    pdf.OpenPdf "C:\Documents\sample.pdf"

    ' Initialize outlines
    pdf.AddOutlines

    ' Create chapter outline
    Dim chapterDefaults As Dictionary
    Set chapterDefaults = New Dictionary
    chapterDefaults.Add "/Title", pdfValue.NewValue("Chapter 1: Introduction")

    Dim dest As pdfValue
    Set dest = pdf.NewDestination(0, PDF_FIT.PDF_FIT)
    chapterDefaults.Add "/Dest", dest

    Dim chapter1 As pdfValue
    Set chapter1 = pdf.NewOutlineItem(pdf.Outlines, chapterDefaults)
    pdf.AddOutlineItem pdf.Outlines, chapter1

    pdf.SavePdfAs "C:\Documents\sample_with_outlines.pdf"
End Sub
```

## Notes and Limitations

- Only supports PDF documents that fit entirely in memory
- Text extraction functionality is limited and currently a work in progress
- Best suited for PDF manipulation rather than creation of complex content
- Object caching is used to maintain performance and avoid Office applications freezing

---
