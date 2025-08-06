# pdfValue Class Documentation

A VBA class representing any value stored in a PDF file, supporting all PDF data types as defined in PDF specification 1.7, Section 3.2.

## Table of Contents
- [Overview](#overview)
- [Enumerations](#enumerations)
- [Properties](#properties)
- [Static Methods](#static-methods)
- [Instance Methods](#instance-methods)
- [Usage Examples](#usage-examples)

## Overview

The `pdfValue` class is the fundamental building block for representing PDF data structures. It provides a unified interface for handling all PDF value types including primitives (null, boolean, numbers, strings, names), collections (arrays, dictionaries), and complex objects (streams, references, objects).

### Key Features
- Supports all 9 basic PDF data types plus references and objects
- Automatic type conversion and validation
- Interned name objects for memory efficiency and fast comparisons
- Bidirectional object-reference relationships
- Singleton instances for common values
- Convenient array creation with `NewArray()` function
- Automatic type conversion with `ConvertToPdfValue()` helper
- Deep cloning with `Clone()` function
- Enhanced `Add()` method with ParamArray support

## Enumerations

### PDF_ValueType
Defines the supported PDF data types as specified in PDF specification 1.7, Section 3.2.

```vba
Public Enum PDF_ValueType
    PDF_Null = 0          ' PDF null object
    PDF_Boolean           ' PDF boolean (true/false)
    PDF_Integer           ' PDF integer number
    PDF_Real              ' PDF real number
    PDF_String            ' PDF string literal
    PDF_Name              ' PDF name object
    PDF_Array             ' PDF array
    PDF_Dictionary        ' PDF dictionary
    PDF_Stream            ' PDF stream object
    PDF_Reference         ' PDF indirect object reference
    PDF_Object            ' PDF indirect object
End Enum
```

### ValueFlags
Defines flags for value representation options.

```vba
Public Enum ValueFlags
    flgNone = 0           ' No special flags
    flgBinary = 1         ' Indicate string represents binary (not textual) data
End Enum
```

## Properties

### ID Property
Gets or sets the object identifier for indirect objects (PDF spec 7.3.10).

```vba
Public Property Let ID(ByVal newId As Long)
Public Property Get ID() As Long
```

**Parameters:**
- `newId` (Long): Object identifier. Range: 1 to 2,147,483,647 for valid objects, 0 for direct objects.

**Returns:**
- `Long`: Current object identifier.

**Usage:**
```vba
Dim obj As New pdfValue
obj.ValueType = PDF_Object
obj.ID = 5  ' Automatically updates associated reference object
Debug.Print obj.ID  ' Outputs: 5
```

### Generation Property
Gets or sets the generation number for indirect objects (PDF spec 7.3.10).

```vba
Public Property Let Generation(ByVal newGeneration As Long)
Public Property Get Generation() As Long
```

**Parameters:**
- `newGeneration` (Long): Generation number. Range: 0 to 65,535, typically 0 for new objects.

**Returns:**
- `Long`: Current generation number.

**Usage:**
```vba
Dim obj As New pdfValue
obj.ValueType = PDF_Object
obj.Generation = 1  ' Automatically updates associated reference object
Debug.Print obj.Generation  ' Outputs: 1
```

### Value Property
The actual value content, type depends on ValueType.

```vba
Public Value As Variant
```

**Content Types:**
- `Boolean` for PDF_Boolean
- `Long` for PDF_Integer
- `Double` for PDF_Real
- `String` for PDF_String/PDF_Name
- `Collection` for PDF_Array
- `Dictionary` for PDF_Dictionary
- `pdfValue` for PDF_Object/PDF_Reference
- `Null` for PDF_Null

### ValueType Property
Specifies the PDF data type of this value.

```vba
Public ValueType As PDF_ValueType
```

**Usage:**
```vba
Dim val As New pdfValue
val.ValueType = PDF_String
val.Value = "Hello World"
```

### Flags Property
Specifies representation flags for the value.

```vba
Public Flags As ValueFlags
```

**Usage:**
```vba
Dim strVal As pdfValue
Set strVal = pdfValue.NewValue(PDF_String, "binary data")
strVal.Flags = strVal.Flags Or ValueFlags.flgBinary  ' Force hex representation
```

### ReferenceObj Property (Read-Only)
Returns reference object for this pdfValue if it represents a PDF object or reference.

```vba
Public Property Get ReferenceObj() As pdfValue
```

**Returns:**
- `pdfValue`: Reference object for PDF_Object types, self for PDF_Reference types, Nothing for other types.

**Usage:**
```vba
Dim obj As New pdfValue
obj.ValueType = PDF_Object
obj.ID = 5
obj.Generation = 0

Dim ref1 As pdfValue, ref2 As pdfValue
Set ref1 = obj.ReferenceObj
Set ref2 = obj.ReferenceObj
Debug.Print (ref1 Is ref2)  ' Outputs: True (same instance)
```

### Count Property (Read-Only)
Returns the count of elements based on value type.

```vba
Public Property Get Count() As Long
```

**Returns:**
- `Long`: Element count (0 for PDF_Null, 1 for scalars, actual count for collections).

**Usage:**
```vba
Dim arr As pdfValue
Set arr = pdfValue.NewValue(PDF_Array)
arr.Add pdfValue.NewValue(PDF_Integer, 1)
arr.Add pdfValue.NewValue(PDF_Integer, 2)
Debug.Print arr.Count ' Outputs: 2
```

## Static Methods

### NullValue Function
Returns PDF null object instance as defined in PDF specification 7.3.9.

```vba
Public Function NullValue() As pdfValue
```

**Returns:**
- `pdfValue`: Singleton null value instance.

**Usage:**
```vba
Dim nullVal As pdfValue
Set nullVal = pdfValue.NullValue()
Debug.Print nullVal.ToString() ' Outputs: "null"
Debug.Print (pdfValue.NullValue() Is pdfValue.NullValue()) ' Outputs: True
```

### TrueValue Function
Returns a PDF boolean object with value True.

```vba
Public Function TrueValue() As pdfValue
```

**Returns:**
- `pdfValue`: Singleton True boolean instance.

**Usage:**
```vba
Dim trueVal As pdfValue
Set trueVal = pdfValue.TrueValue()
Debug.Print trueVal.ToString() ' Outputs: "true"
```

### FalseValue Function
Returns a PDF boolean object with value False.

```vba
Public Function FalseValue() As pdfValue
```

**Returns:**
- `pdfValue`: Singleton False boolean instance.

**Usage:**
```vba
Dim falseVal As pdfValue
Set falseVal = pdfValue.FalseValue()
Debug.Print falseVal.ToString() ' Outputs: "false"
```

### NewValue Function
Creates a PDF value of the specified type.

```vba
Public Function NewValue(ByVal valueType As PDF_ValueType, _
    Optional ByVal value As Variant, _
    Optional ByVal objId As Long = 0, _
    Optional ByVal generation As Long = 0) As pdfValue
```

**Parameters:**
- `valueType` (PDF_ValueType): The type of PDF value to create.
- `value` (Variant, Optional): Primary value content (required for most types).
- `objId` (Long, Optional): Object identifier for PDF_Object/PDF_Reference types.
- `generation` (Long, Optional): Generation number, defaults to 0.

**Returns:**
- `pdfValue`: New pdfValue instance of specified type.

**Usage Examples:**
```vba
' Create null value
Set nullVal = pdfValue.NewValue(PDF_Null)

' Create boolean values
Set trueVal = pdfValue.NewValue(PDF_Boolean, True)
Set falseVal = pdfValue.NewValue(PDF_Boolean, False)

' Create numeric values
Set intVal = pdfValue.NewValue(PDF_Integer, 42)
Set realVal = pdfValue.NewValue(PDF_Real, 3.14159)

' Create string values
Set strVal = pdfValue.NewValue(PDF_String, "Hello World")

' Create name values
Set nameVal = pdfValue.NewValue(PDF_Name, "Type")
Set nameVal = pdfValue.NewValue(PDF_Name, "/Type")  ' Leading slash optional

' Create collection values
Set arrVal = pdfValue.NewValue(PDF_Array)
Set dictVal = pdfValue.NewValue(PDF_Dictionary)

' Create reference values
Set refVal = pdfValue.NewValue(PDF_Reference, objId:=5)

' Create object values
Dim contentDict As pdfValue
Set contentDict = pdfValue.NewValue(PDF_Dictionary)
Set objVal = pdfValue.NewValue(PDF_Object, contentDict, 5, 0)
```

### NewArray Function
Creates a PDF array with variable number of items using ParamArray for flexible initialization.

```vba
Public Function NewArray(ParamArray items() As Variant) As pdfValue
```

**Parameters:**
- `items` (ParamArray Variant): Variable number of items to add to the array.

**Type Conversion Rules:**
- Numeric values become PDF_Integer or PDF_Real based on decimal content
- String values become PDF_String (literal strings) or PDF_Name (if prefixed with "/")
- Boolean values become PDF_Boolean
- Existing pdfValue objects are added directly
- VBA Arrays are expanded into individual elements
- Other objects trigger an error

**Returns:**
- `pdfValue`: A new pdfValue with ValueType = PDF_Array containing all specified items.

**Usage Examples:**
```vba
' Create coordinate array for PDF rectangle [0 0 612 792]
Dim mediaBox As pdfValue
Set mediaBox = pdfValue.NewArray(0, 0, 612, 792)

' Create mixed content array [3.4 0 0 52 (some string) (another string) 4.2 false true]
Dim mixedArray As pdfValue
Set mixedArray = pdfValue.NewArray(3.4, 0, 0, 52, "some string", "another string", 4.2, False, True)

' Create color array for RGB [1.0 0.5 0.0]
Dim rgbColor As pdfValue
Set rgbColor = pdfValue.NewArray(1.0, 0.5, 0.0)

' Create array with PDF names [/Type /Page /Parent 3 0 R]
Dim pageRef As pdfValue
Set pageRef = pdfValue.NewValue(PDF_Reference, , 3, 0)
Dim pageArray As pdfValue
Set pageArray = pdfValue.NewArray("/Type", "/Page", "/Parent", pageRef)

' Create transformation matrix [1 0 0 1 100 200]
Dim transform As pdfValue
Set transform = pdfValue.NewArray(1, 0, 0, 1, 100, 200)

' Empty array creation
Dim emptyArray As pdfValue
Set emptyArray = pdfValue.NewArray()

' Create from VBA Array() function
Dim coords As pdfValue
Set coords = pdfValue.NewArray(Array(0, 0, 612, 792))

' Create from computed array
Dim values As Variant
values = Array(1.0, 2.0, 3.0, 4.0)
Dim numArray As pdfValue
Set numArray = pdfValue.NewArray(values)

' Create from worksheet range values
Dim rangeValues As Variant
rangeValues = Range("A1:D1").Value  ' Returns 1D array
Dim rowArray As pdfValue
Set rowArray = pdfValue.NewArray(rangeValues)
```

### NewTypedDictionary Function
Creates a PDF dictionary with specified /Type and optional /Subtype entries.

```vba
Public Function NewTypedDictionary(ByVal typeName As String, _
    Optional ByVal subTypeName As String = vbNullString) As pdfValue
```

**Parameters:**
- `typeName` (String): PDF object type name for /Type entry (slash added automatically).
- `subTypeName` (String, Optional): PDF object subtype name for /Subtype entry.

**Returns:**
- `pdfValue`: New dictionary with /Type and optional /Subtype entries.

**Usage Examples:**
```vba
' Create a page dictionary
Dim pageDict As pdfValue
Set pageDict = pdfValue.NewTypedDictionary("Page")

' Create a font dictionary with subtype
Dim fontDict As pdfValue
Set fontDict = pdfValue.NewTypedDictionary("Font", "Type1")

' Create an annotation dictionary
Dim annotDict As pdfValue
Set annotDict = pdfValue.NewTypedDictionary("Annot", "Text")
```

### GetInternedName Function
Returns an interned PDF name object to avoid duplicates and enable fast comparisons.

```vba
Public Function GetInternedName(ByVal nameValue As String) As pdfValue
```

**Parameters:**
- `nameValue` (String): Name content with or without leading slash.

**Returns:**
- `pdfValue`: Cached name instance (same instance for identical names).

**Usage:**
```vba
Dim typeVal1 As pdfValue, typeVal2 As pdfValue
Set typeVal1 = pdfValue.GetInternedName("Type")
Set typeVal2 = pdfValue.GetInternedName("/Type")
Debug.Print (typeVal1 Is typeVal2) ' Outputs: True (same instance)

' Common usage in PDF dictionaries
dict.SetDictionaryValue "/Type", pdfValue.GetInternedName("Page")
```

## Instance Methods

### Exists Function
Safely checks if a key exists in dictionary values.

```vba
Public Function Exists(ByVal key As String) As Boolean
```

**Parameters:**
- `key` (String): Dictionary key to check (slash added automatically if missing).

**Returns:**
- `Boolean`: True if key exists, False otherwise or if not a dictionary type.

**Usage:**
```vba
Dim dict As pdfValue
Set dict = pdfValue.NewValue(PDF_Dictionary)
dict.SetDictionaryValue "/Type", pdfValue.GetInternedName("Page")

Debug.Print dict.Exists("Type")    ' Outputs: True
Debug.Print dict.Exists("/Type")   ' Outputs: True
Debug.Print dict.Exists("Missing") ' Outputs: False

' Works with PDF objects containing dictionaries
Dim obj As New pdfValue
obj.ValueType = PDF_Object
Set obj.Value = dict
Debug.Print obj.Exists("Type")    ' Outputs: True
```

### GetDictionaryValue Function
Gets values from dictionary types including nested dictionaries in PDF_Object types.

```vba
Public Function GetDictionaryValue(ByVal key As String) As Variant
```

**Parameters:**
- `key` (String): Dictionary key to look up (slash added automatically if missing).

**Returns:**
- `Variant`: Value associated with key, or Nothing if not found.

**Usage:**
```vba
Dim dict As pdfValue
Set dict = pdfValue.NewValue(PDF_Dictionary)
dict.SetDictionaryValue "/Type", pdfValue.GetInternedName("Page")

Dim typeVal As pdfValue
Set typeVal = dict.GetDictionaryValue("Type")
Debug.Print typeVal.ToString() ' Outputs: "/Page"
```

### SetDictionaryValue Subroutine
Sets values in dictionary types including nested dictionaries in PDF_Object types.

```vba
Public Sub SetDictionaryValue(ByVal key As String, ByRef item As Variant)
```

**Parameters:**
- `key` (String): Dictionary key (slash added automatically if missing).
- `item` (Variant): Value to associate with key.

**Usage:**
```vba
Dim dict As pdfValue
Set dict = pdfValue.NewValue(PDF_Dictionary)
dict.SetDictionaryValue "/Type", pdfValue.GetInternedName("Page")
dict.SetDictionaryValue "/MediaBox", pdfValue.NewValue(PDF_Array)
```

### Add Subroutine
Adds one or more items to the end of a PDF array value with automatic type conversion.

```vba
Public Sub Add(ParamArray items() As Variant)
```

**Parameters:**
- `items` (ParamArray Variant): Variable number of items to add to the array.

**Type Conversion:**
- Non-pdfValue items are automatically converted using `ConvertToPdfValue()`
- Supports mixed types in a single call

**Usage:**
```vba
Dim arr As pdfValue
Set arr = pdfValue.NewValue(PDF_Array)

' Traditional single item addition
arr.Add pdfValue.NewValue(PDF_Integer, 42)

' Multiple items at once with automatic conversion
arr.Add 1, 2, 3, "text", True, "/Name"
' Equivalent to multiple individual Add calls:
' arr.Add pdfValue.NewValue(PDF_Integer, 1)
' arr.Add pdfValue.NewValue(PDF_Integer, 2)
' arr.Add pdfValue.NewValue(PDF_Integer, 3)
' arr.Add pdfValue.NewValue(PDF_String, "text")
' arr.Add pdfValue.TrueValue()
' arr.Add pdfValue.GetInternedName("/Name")

' Build array incrementally
Dim coords As pdfValue
Set coords = pdfValue.NewValue(PDF_Array)
coords.Add 0, 0    ' Add first two coordinates
coords.Add 612, 792    ' Add remaining coordinates

' Build arrays incrementally with multiple items
Dim mediaBox As pdfValue
Set mediaBox = pdfValue.NewValue(PDF_Array)
mediaBox.Add 0, 0    ' Add first two coordinates
mediaBox.Add 612, 792    ' Add remaining coordinates
```

### GetArrayValue Function
Gets a value from a PDF array by 1-based index.

```vba
Public Function GetArrayValue(ByVal index As Long) As Variant
```

**Parameters:**
- `index` (Long): 1-based index of array element to retrieve.

**Returns:**
- `Variant`: Value at specified index, or Nothing if invalid index or not an array.

**Usage:**
```vba
Dim arr As pdfValue
Set arr = pdfValue.NewValue(PDF_Array)
arr.Add pdfValue.NewValue(PDF_Integer, 42)
arr.Add pdfValue.NewValue(PDF_String, "text")

Dim firstItem As pdfValue
Set firstItem = arr.GetArrayValue(1) ' Gets first element (1-based)
Debug.Print firstItem.ToString() ' Outputs: "42"

Dim secondItem As pdfValue
Set secondItem = arr.GetArrayValue(2) ' Gets second element
Debug.Print secondItem.ToString() ' Outputs: "(text)"
```

### Clone Function
Performs a deep copy of a pdfValue object tree.

```vba
Public Function Clone(Optional ByRef objRefCache As Dictionary = Nothing) As pdfValue
```

**Parameters:**
- `objRefCache` (Dictionary, Optional): Cache for object references to maintain consistency.

**Returns:**
- `pdfValue`: Deep copy of the original pdfValue and its entire object tree.

**Usage:**
```vba
' Clone a simple value
Dim original As pdfValue
Set original = pdfValue.NewValue(PDF_String, "Hello World")
Dim copy As pdfValue
Set copy = original.Clone()

' Clone a complex structure
Dim originalDict As pdfValue
Set originalDict = pdfValue.NewTypedDictionary("Page")
originalDict.SetDictionaryValue "/MediaBox", pdfValue.NewArray(0, 0, 612, 792)

Dim clonedDict As pdfValue
Set clonedDict = originalDict.Clone()

' Changes to clone don't affect original
clonedDict.SetDictionaryValue "/Rotate", pdfValue.NewValue(PDF_Integer, 90)
' originalDict remains unchanged

' Clone with object reference cache for consistency
Dim objCache As New Dictionary
Dim obj1 As pdfValue: Set obj1 = someComplexObject.Clone(objCache)
Dim obj2 As pdfValue: Set obj2 = anotherObject.Clone(objCache)
' Shared references between obj1 and obj2 are preserved
```

### ToString Function
Converts the PDF value to its string representation for PDF serialization.

```vba
Public Function ToString() As String
```

**Returns:**
- `String`: PDF-formatted representation suitable for PDF content streams.

**Usage:**
```vba
Debug.Print myValue.ToString()
fileContent = fileContent & myValue.ToString()
```

**Output Examples:**
- PDF_Null: `"null"`
- PDF_Boolean: `"true"` or `"false"`
- PDF_Integer: `"42"`
- PDF_Real: `"3.14159"`
- PDF_String: `"(Hello World)"` or `<48656C6C6F20576F726C64>` (hex)
- PDF_Name: `"/Type"`
- PDF_Array: `"[1 2 3]"`
- PDF_Dictionary: `"<< /Type /Page >>"`
- PDF_Reference: `"5 0 R"`

### Serialize Function
Converts the PDF value to its binary representation as stored in PDF files.

```vba
Public Function Serialize() As Byte()
```

**Returns:**
- `Byte()`: Binary representation suitable for writing to PDF files.

**Usage:**
```vba
Dim pdfBytes() As Byte
pdfBytes = myValue.Serialize()
' Write pdfBytes to file
```

## Helper Methods

### ConvertToPdfValue Function
Internal helper method to convert various VBA types to appropriate pdfValue objects.

```vba
Function ConvertToPdfValue(ByRef item As Variant) As pdfValue
```

**Type Conversion Rules:**
- Integers and Longs become PDF_Integer
- Singles and Doubles become PDF_Real
- Strings starting with "/" become PDF_Name (with slash preserved)
- Other strings become PDF_String
- Boolean values become PDF_Boolean (using shared True/False instances)
- Existing pdfValue objects pass through unchanged
- Dictionary objects become PDF_Dictionary
- Collection objects become PDF_Array
- Nothing or Null becomes PDF_Null
- Other object types raise an error

**Parameters:**
- `item` (Variant): The value to convert to a pdfValue object.

**Returns:**
- `pdfValue`: A pdfValue object representing the input value.

**Usage:**
```vba
' Internal use - called automatically by Add() and NewArray()
Dim converted As pdfValue
Set converted = ConvertToPdfValue(42)        ' Returns PDF_Integer
Set converted = ConvertToPdfValue(3.14)      ' Returns PDF_Real
Set converted = ConvertToPdfValue("text")    ' Returns PDF_String
Set converted = ConvertToPdfValue("/Name")   ' Returns PDF_Name
Set converted = ConvertToPdfValue(True)      ' Returns PDF_Boolean
```

## Usage Examples

### Creating Basic Values
```vba
' Create primitive values
Dim nullVal As pdfValue: Set nullVal = pdfValue.NullValue()
Dim boolVal As pdfValue: Set boolVal = pdfValue.NewValue(PDF_Boolean, True)
Dim intVal As pdfValue: Set intVal = pdfValue.NewValue(PDF_Integer, 42)
Dim realVal As pdfValue: Set realVal = pdfValue.NewValue(PDF_Real, 3.14159)
Dim strVal As pdfValue: Set strVal = pdfValue.NewValue(PDF_String, "Hello")
Dim nameVal As pdfValue: Set nameVal = pdfValue.GetInternedName("Type")
```

### Working with Arrays - Traditional Approach
```vba
' Create and populate array traditionally
Dim mediaBox As pdfValue
Set mediaBox = pdfValue.NewValue(PDF_Array)
mediaBox.Add pdfValue.NewValue(PDF_Integer, 0)
mediaBox.Add pdfValue.NewValue(PDF_Integer, 0)
mediaBox.Add pdfValue.NewValue(PDF_Integer, 612)
mediaBox.Add pdfValue.NewValue(PDF_Integer, 792)

Debug.Print mediaBox.ToString() ' Outputs: "[0 0 612 792]"
Debug.Print mediaBox.Count    ' Outputs: 4
```

### Working with Arrays - Convenient Approach
```vba
' Create arrays with NewArray() - much more concise
Dim mediaBox As pdfValue
Set mediaBox = pdfValue.NewArray(0, 0, 612, 792)

' Create mixed content arrays easily
Dim mixedArray As pdfValue
Set mixedArray = pdfValue.NewArray(3.14, "text", True, "/Name", 42)

' Add multiple items at once with enhanced Add()
Dim coords As pdfValue
Set coords = pdfValue.NewValue(PDF_Array)
coords.Add 100, 200, 300, 400  ' Automatic type conversion

' Access array elements
Dim width As pdfValue
Set width = mediaBox.GetArrayValue(3) ' Gets third element (612)
```

### Advanced Array Creation Examples
```vba
' Create coordinate arrays with automatic type detection
Dim mediaBox As pdfValue
Set mediaBox = pdfValue.NewArray(0, 0, 612, 792)
Debug.Print mediaBox.ToString() ' Outputs: "[0 0 612 792]"

' Create mixed-type arrays
Dim mixedArray As pdfValue
Set mixedArray = pdfValue.NewArray(42, 3.14, "text", "/Name", True)
Debug.Print mixedArray.ToString() ' Outputs: "[42 3.14 (text) /Name true]"

' Create transformation matrices
Dim ctm As pdfValue
Set ctm = pdfValue.NewArray(1.5, 0, 0, 1.5, 100, 200)

' Create color arrays
Dim rgbRed As pdfValue: Set rgbRed = pdfValue.NewArray(1.0, 0.0, 0.0)
Dim cmykBlack As pdfValue: Set cmykBlack = pdfValue.NewArray(0, 0, 0, 1.0)

' Create from VBA arrays
Dim coords As Variant: coords = Array(72, 144, 216, 288)
Dim coordArray As pdfValue: Set coordArray = pdfValue.NewArray(coords)

' Build arrays incrementally with multiple items
Dim pathArray As pdfValue
Set pathArray = pdfValue.NewValue(PDF_Array)
pathArray.Add 100, 100, "m"    ' Move to 100,100
pathArray.Add 200, 100, "l"    ' Line to 200,100  
pathArray.Add 200, 200, "l"    ' Line to 200,200
pathArray.Add "h", "S"    ' Close path and stroke

' Create transformation matrices easily
Dim identityMatrix As pdfValue
Set identityMatrix = pdfValue.NewArray(1, 0, 0, 1, 0, 0)

' Create color spaces
Dim rgbColor As pdfValue
Set rgbColor = pdfValue.NewArray(1.0, 0.5, 0.0)  ' Orange color

Dim cmykColor As pdfValue
Set cmykColor = pdfValue.NewArray(0, 0.5, 1.0, 0)  ' Orange in CMYK

' Create complex nested structures
Dim fontDict As pdfValue
Set fontDict = pdfValue.NewTypedDictionary("Font", "Type1")
fontDict.SetDictionaryValue "/BaseFont", pdfValue.GetInternedName("Helvetica")

Dim resourceDict As pdfValue
Set resourceDict = pdfValue.NewValue(PDF_Dictionary)
resourceDict.SetDictionaryValue "/Font", fontDict

' Create arrays from Excel ranges
Dim rangeData As Variant
rangeData = Range("A1:D1").Value
Dim excelArray As pdfValue
Set excelArray = pdfValue.NewArray(rangeData)

' Create arrays from computed values
Dim i As Long
Dim computedValues As pdfValue
Set computedValues = pdfValue.NewValue(PDF_Array)
For i = 1 To 10
    computedValues.Add i * 2, i * 3  ' Add pairs of computed values
Next i
```

### Working with Dictionaries
```vba
' Create page dictionary
Dim pageDict As pdfValue
Set pageDict = pdfValue.NewTypedDictionary("Page")

' Add properties
pageDict.SetDictionaryValue "/Parent", pdfValue.NewValue(PDF_Reference, objId:=3)
pageDict.SetDictionaryValue "/MediaBox", pdfValue.NewArray(0, 0, 612, 792)
pageDict.SetDictionaryValue "/Rotate", pdfValue.NewValue(PDF_Integer, 0)

' Check for keys
If pageDict.Exists("MediaBox") Then
    Dim box As pdfValue
    Set box = pageDict.GetDictionaryValue("MediaBox")
    Debug.Print "MediaBox: " & box.ToString()
End If

Debug.Print pageDict.ToString()
' Outputs: "<< /Type /Page /Parent 3 0 R /MediaBox [0 0 612 792] /Rotate 0 >>"
```

### Creating PDF Objects
```vba
' Create a PDF object with dictionary content
Dim pageObj As pdfValue
Set pageObj = pdfValue.NewValue(PDF_Object, pageDict, 5, 0)

Debug.Print pageObj.ID    ' Outputs: 5
Debug.Print pageObj.Generation ' Outputs: 0

' Get reference to this object
Dim pageRef As pdfValue
Set pageRef = pageObj.ReferenceObj
Debug.Print pageRef.ToString() ' Outputs: "5 0 R"
```

### Working with References
```vba
' Create reference to existing object
Dim fontRef As pdfValue
Set fontRef = pdfValue.NewValue(PDF_Reference, objId:=10, generation:=0)

' Add reference to page dictionary
pageDict.SetDictionaryValue "/Font", fontRef

' References can link to actual objects
Dim fontObj As pdfValue
Set fontObj = pdfValue.NewValue(PDF_Object, fontDict, 10, 0)
Set fontRef.Value = fontObj ' Link reference to actual object
```

### Deep Copying Objects
```vba
' Clone complex structures for independent modification
Dim originalPage As pdfValue
Set originalPage = pdfValue.NewTypedDictionary("Page")
originalPage.SetDictionaryValue "/MediaBox", pdfValue.NewArray(0, 0, 612, 792)
originalPage.SetDictionaryValue "/Rotate", pdfValue.NewValue(PDF_Integer, 0)

' Create independent copy
Dim clonedPage As pdfValue
Set clonedPage = originalPage.Clone()

' Modify clone without affecting original
clonedPage.SetDictionaryValue "/Rotate", pdfValue.NewValue(PDF_Integer, 90)
clonedPage.GetDictionaryValue("/MediaBox").Add 100  ' Add margin

' Original remains unchanged
Debug.Print originalPage.GetDictionaryValue("/Rotate").ToString() ' Still "0"

' Clone simple values
Dim original As pdfValue
Set original = pdfValue.NewValue(PDF_String, "Original Text")
Dim copy As pdfValue
Set copy = original.Clone()

' Modify copy without affecting original
copy.Value = "Modified Text"
Debug.Print original.ToString() ' Outputs: "(Original Text)"
Debug.Print copy.ToString()    ' Outputs: "(Modified Text)"

' Clone complex structures
Dim complexDict As pdfValue
Set complexDict = pdfValue.NewTypedDictionary("Page")
complexDict.SetDictionaryValue "/MediaBox", pdfValue.NewArray(0, 0, 612, 792)
complexDict.SetDictionaryValue "/Resources", pdfValue.NewTypedDictionary("Resources")

Dim clonedDict As pdfValue
Set clonedDict = complexDict.Clone()

' Clone with reference cache for document-wide consistency
Dim objCache As New Dictionary
Dim page1 As pdfValue: Set page1 = pageTemplate.Clone(objCache)
Dim page2 As pdfValue: Set page2 = pageTemplate.Clone(objCache)
' Shared references between pages are preserved
```

### Memory Efficiency with Interned Names
```vba
' Use interned names for frequently used PDF names
Dim type1 As pdfValue: Set type1 = pdfValue.GetInternedName("Type")
Dim type2 As pdfValue: Set type2 = pdfValue.GetInternedName("Type")
Dim page1 As pdfValue: Set page1 = pdfValue.GetInternedName("Page")
Dim page2 As pdfValue: Set page2 = pdfValue.GetInternedName("Page")

' Same names return identical instances
Debug.Print (type1 Is type2) ' Outputs: True
Debug.Print (page1 Is page2) ' Outputs: True
Debug.Print (type1 Is page1) ' Outputs: False

' Fast comparison using Is operator instead of string comparison
If dict.GetDictionaryValue("/Type") Is pdfValue.GetInternedName("Page") Then
    Debug.Print "This is a page object"
End If
```

## See Also
- [pdfDocument Class](pdfDocument.md) - Main document management class
- [PDF Specification 1.7](https://www.adobe.com/content/dam/acom/en/devnet/pdf/pdfs/PDF32000_2008.pdf) - Official PDF specification
- [VBA Dictionary Object](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object) - Used for pdfDictionary storage
- [VBA Collection Object](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/collection-object) - Used for pdfArray storage
