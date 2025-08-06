VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pdfValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ====
' pdfValue.cls - PDF Value Representation
' ====
Option Explicit


' Represents a PDF value supporting all PDF data types as defined in PDF specification 1.7, Section 3.2.
Public Enum PDF_ValueType
    PDF_Null = 0
    PDF_Boolean
    PDF_Integer
    PDF_Real
    PDF_Name
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
    
    ' not a PDF_ValueType per se, but logically one in a content stream
    PDF_Operator = 128  ' high bit of low byte set flags operator, high byte is PdfTextOp
End Enum

' Properties of each unique pdfValue instance

' used to help encode certain information about how to represent
Public Enum ValueFlags
    flgNone = 0
    flgBinary = 1   ' currently only applies to String values, forces <##> format if flgBinary is set instead of (s)
End Enum
Public flags As ValueFlags

' Object identifier for indirect objects (PDF spec 7.3.10).
' Range: 1 to 2,147,483,647 for valid object numbers, 0 for direct objects.
Private m_id As Long

' Generation number for indirect objects (PDF spec 7.3.10).
' Range: 0 to 65,535, typically 0 for new objects.
Private m_generation As Long

' The actual value content, type depends on ValueType.
' Can contain: Boolean, Double, String, Collection, Dictionary, or formatted reference string.
Public value As Variant

' Specifies the PDF data type of this value.
' Determines how Value should be interpreted and serialized.
Public valueType As PDF_ValueType

' Private reference to the interned reference object for this pdfValue when it represents a PDF object
' to help maintain id across indirect references, PDF_Object provides a PDF_Reference object to be used by others
' Note: reference objects are per instance so external to pdfValue object must be used to intern
' as identified by object id and generation which can conflict with distinct objects across different pdf documents
Private m_referenceObj As pdfValue


' Properties and Functions used by the predeclared pdfValue instance
' allowing static like access for convenience of creating new pdfValue instances

' Private dictionary for interned name objects to avoid duplicates and enable fast comparisons
Private m_InternedNames As Dictionary


' Class initialization
Private Sub Class_Initialize()
    '
End Sub

' Class cleanup
Private Sub Class_Terminate()
    Set value = Nothing
    Set m_referenceObj = Nothing
    Set m_InternedNames = Nothing
End Sub

' Synchronizes the ID property with any associated reference object to maintain consistency.
' When this pdfValue represents a PDF object, updates the reference object's value to match.
'
' Parameters:
'   newId (Long) - The new object identifier value.
'    Range: 1 to 2,147,483,647 for valid object numbers, 0 for direct objects.
'
' Usage:
'   Dim obj As New pdfValue
'   obj.ValueType = PDF_Object
'   obj.ID = 5  ' Automatically updates associated reference object
Public Property Let id(ByVal newId As Long)
    m_id = newId
    If valueType = PDF_Object Then
        If Not m_referenceObj Is Nothing Then
            m_referenceObj.id = m_id
        End If
    End If
End Property

' Returns the object identifier for this PDF value.
'
' Returns:
'   Long - The object identifier (0 for direct objects, >0 for indirect objects).
'
' Usage:
'   Debug.Print obj.ID  ' Outputs the object number
Public Property Get id() As Long
    id = m_id
End Property

' Synchronizes the generation property with any associated reference object to maintain consistency.
' When this pdfValue represents a PDF object, updates the reference object's value to match.
'
' Parameters:
'   newGeneration (Long) - The new generation number value.
'    Range: 0 to 65,535, typically 0 for new objects.
'
' Usage:
'   Dim obj As New pdfValue
'   obj.ValueType = PDF_Object
'   obj.Generation = 1  ' Automatically updates associated reference object
Public Property Let generation(ByVal newGeneration As Long)
    m_generation = newGeneration
    If valueType = PDF_Object Then
        If Not m_referenceObj Is Nothing Then
            m_referenceObj.generation = m_generation
        End If
    End If
End Property

' Returns the generation number for this PDF value.
'
' Returns:
'   Long - The generation number (typically 0 for new objects).
'
' Usage:
'   Debug.Print obj.Generation  ' Outputs the generation number
Public Property Get generation() As Long
    generation = m_generation
End Property

' Internal method to set the reference object for this pdfValue.
' Used by the referenceObj property getter to establish the bidirectional relationship.
'
' Parameters:
'   refObj (pdfValue) - The reference object to associate with this pdfValue.
'
' Usage:
'   ' Internal use only - called automatically by referenceObj property
Friend Property Set referenceObj(ByRef refObj As pdfValue)
    Set m_referenceObj = refObj
End Property

' Returns reference object for this pdfValue if it represents a PDF object or reference.
' Returns Nothing if this pdfValue is not a PDF object or reference type.
'
' For PDF_Object types: Creates a PDF_Reference that points to this object.
' For PDF_Reference types: Returns self.
' For all other types: Returns Nothing.
'
' Returns:
'   pdfValue - A reference object, or Nothing if not applicable.
'   Note: an external object should be used to ensure the same instance
'         is returned for identical object ID and generation combinations.
'
' Usage:
'   Dim obj As New pdfValue
'   obj.ValueType = PDF_Object
'   obj.ID = 5
'   obj.Generation = 0
'
'   Dim ref1 As pdfValue, ref2 As pdfValue
'   Set ref1 = obj.referenceObj
'   Set ref2 = obj.referenceObj
'   Debug.Print (ref1 Is ref2)  ' Outputs: True (same instance)
Public Property Get referenceObj() As pdfValue
    Select Case valueType
        Case PDF_Object
            ' Create reference object if not already exists
            If (m_referenceObj Is Nothing) And (Me.id <> 0) Then
                ' Create new reference object
                Dim newRef As New pdfValue
                Set m_referenceObj = pdfValue.NewValue(PDF_Reference, Me, Me.id, Me.generation)
            End If
            
            ' Always return same object so all references updated if id changes
            Set referenceObj = m_referenceObj
            
        Case PDF_Reference
            Set referenceObj = Me
            
        Case Else
            ' Not a reference or object type
            Set referenceObj = Nothing
    End Select
End Property

' Safely checks if a key exists in the dictionary value of this pdfValue.
' Works for both direct PDF_Dictionary values and PDF_Object values that contain dictionaries.
' Returns False if this pdfValue is not a dictionary type or if the key doesn't exist.
'
' For PDF_Dictionary types: Directly checks the dictionary stored in Value.
' For PDF_Object types: Checks the dictionary stored in the nested pdfValue's Value property.
' For all other types: Returns False.
'
' Parameters:
'   key (String) - The dictionary key to check for existence.
'    Range: Any string that might exist as a key.
'    Expected: PDF name format (e.g., "/Type", "MediaBox").
'    Note: Leading slash is added automatically if not present.
'
' Returns:
'   Boolean - True if the key exists in the dictionary, False otherwise.
'
' Usage:
'   Dim dict As pdfValue
'   Set dict = pdfValue.NewValue(PDF_Dictionary)
'   dict.SetDictionaryValue "/Type", pdfValue.GetInternedName("Page")
'
'   Debug.Print dict.Exists("Type")     ' Outputs: True
'   Debug.Print dict.Exists("/Type")    ' Outputs: True
'   Debug.Print dict.Exists("Missing")  ' Outputs: False
'
'   ' Works with PDF objects containing dictionaries
'   Dim obj As New pdfValue
'   obj.ValueType = PDF_Object
'   Set obj.Value = dict
'   Debug.Print obj.Exists("Type")      ' Outputs: True
Public Function Exists(ByVal key As String) As Boolean
    ' Ensure key starts with forward slash for consistency
    If Left(key, 1) <> "/" Then
        key = "/" & key
    End If
    
    Select Case valueType
        Case PDF_Dictionary
            ' Direct dictionary access
            If typeName(value) = "Dictionary" Then
                Exists = value.Exists(key)
            'Else
            '    Exists = False
            End If
            
        Case PDF_Object
            ' PDF object containing a dictionary
            If typeName(value) = "pdfValue" Then
                Dim innerValue As pdfValue
                Set innerValue = value
                If innerValue.valueType = PDF_Dictionary And typeName(innerValue.value) = "Dictionary" Then
                    Exists = innerValue.value.Exists(key)
                'Else
                '    Exists = False
                End If
            'Else
            '    Exists = False
            End If
            
        Case PDF_Reference
            If typeName(value) = "pdfValue" Then    ' if Not value Is Nothing and its a pdfValue
                Exists = value.Exists(key)
            End If
            
        Case Else
            ' Not a dictionary or object type
            'Exists = False
    End Select
End Function

' Enhanced GetDictionaryValue that works with both PDF_Dictionary and PDF_Object types.
' For PDF_Object types, accesses the dictionary stored in the nested pdfValue.
'
' Parameters:
'   key (String) - The dictionary key to look up.
'
' Returns:
'   Variant - The value associated with the key, or Nothing if key not found or not a dictionary.
Public Function GetDictionaryValue(ByVal key As String) As Variant
    On Error GoTo errHandler
    Set GetDictionaryValue = Nothing

    ' Ensure key starts with forward slash for consistency
    If Left(key, 1) <> "/" Then
        key = "/" & key
    End If
    
    Select Case valueType
        Case PDF_Dictionary
            ' Direct dictionary access - existing functionality
            If typeName(value) = "Dictionary" Then
                If value.Exists(key) Then
                    Dim item As Variant
                    Set item = value(key)
                    If IsObject(item) Then
                        Set GetDictionaryValue = item
                    Else
                        GetDictionaryValue = item
                    End If
                End If
            End If
            
        Case PDF_Object
            ' PDF object containing a dictionary
            If typeName(value) = "pdfValue" Then
                Dim innerValue As pdfValue
                Set innerValue = value
                If innerValue.valueType = PDF_Dictionary And typeName(innerValue.value) = "Dictionary" Then
                    If innerValue.value.Exists(key) Then
                        If IsObject(innerValue.value(key)) Then
                            Set GetDictionaryValue = innerValue.value(key)
                        Else
                            Let GetDictionaryValue = innerValue.value(key)
                        End If
                    End If
                End If
            End If
            
        Case PDF_Reference
            ' if object instance is linked to reference then get value from it
            If typeName(value) = "pdfValue" Then    ' if Not value Is Nothing and its a pdfValue
                If IsObject(value.GetDictionaryValue(key)) Then
                    Set GetDictionaryValue = value.GetDictionaryValue(key)
                Else
                    Let GetDictionaryValue = value.GetDictionaryValue(key)
                End If
            End If
            
        Case Else
            'GetDictionaryValue = Nothing
    End Select
    
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Function

' Sets values in the dictionary.
' For PDF_Object types updates nested PDF_Dictionary value
' for PDF_Dictionary updates Dictionary stored in value
' Ignored for all other types
'
' Parameters:
'   key (String) - The dictionary key.
'   item (Variant) - The value to associate with the key.
Public Sub SetDictionaryValue(ByVal key As String, ByRef item As Variant)
    On Error GoTo errHandler
    
    ' Ensure key starts with forward slash for consistency
    If Left(key, 1) <> "/" Then
        key = "/" & key
    End If
    
    Select Case valueType
        Case PDF_Dictionary
            ' Direct dictionary access
            If typeName(value) = "Dictionary" Then
                If IsObject(item) Then
                    Set value.item(key) = item
                Else
                    value.item(key) = item
                End If
            End If
            
        Case PDF_Object
            ' PDF object containing a dictionary
            If typeName(value) = "pdfValue" Then
                Dim innerValue As pdfValue
                Set innerValue = value
                If innerValue.valueType = PDF_Dictionary And typeName(innerValue.value) = "Dictionary" Then
                    If IsObject(item) Then
                        Set innerValue.value(key) = item
                    Else
                        Let innerValue.value(key) = item
                    End If
                End If
            End If
        
        Case PDF_Reference
            ' if object instance is linked to reference then get value from it
            If typeName(value) = "pdfValue" Then    ' if Not value Is Nothing and its a pdfValue
                value.SetDictionaryValue key, item
            End If
        
        Case Else
            Debug.Print "Warning: unsupported action - SetDictionaryValue on value that is not a dictionary!"
            Stop
    End Select

    Exit Sub
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Sub

' Returns PDF null object instance as defined in PDF specification 7.3.9.
'
' PDF null objects represent missing or undefined values in PDF structures.
' Commonly used in arrays and dictionaries to indicate absent entries.
' The null object is written as the keyword "null" in PDF content streams.
' This method always returns the same instance, if used consistently
' then two null objects can be directly compared obj1 Is obj2.
'
' Returns:
'   pdfValue - A pdfValue instance with ValueType = PDF_Null and Value = Null.
'
' Usage:
'   Dim nullVal As pdfValue
'   Set nullVal = pdfValue.NullValue()
'   Debug.Print nullVal.ToString() ' Outputs: "null"
'   Debug.Print (pdfValue.NullValue() Is pdfValue.NullValue()) ' Outputs: True
Public Function NullValue() As pdfValue
    If Not Me Is pdfValue Then
        Set NullValue = pdfValue.NullValue
        Exit Function
    End If
    
    Static m_NullInstance As pdfValue
    If m_NullInstance Is Nothing Then
        Set m_NullInstance = New pdfValue
        m_NullInstance.valueType = PDF_Null
        m_NullInstance.value = Null
        'm_NullInstance.m_id = 0
        'm_NullInstance.m_generation = 0
    End If
    Set NullValue = m_NullInstance
End Function

' Returns a PDF boolean object with value True as defined in PDF specification 7.3.2.
'
' PDF boolean objects represent logical true/false values using keywords "true" and "false".
' Used in PDF dictionaries for flags, options, and conditional settings.
' Boolean values are case-sensitive and must be lowercase in PDF syntax.
' This method always returns the same True boolean instance, if used consistently
' then two true objects can be directly compared obj1 Is obj2.
'
' Returns:
'   pdfValue - A pdfValue instance with ValueType = PDF_Boolean and Value = True.
'
' Usage:
'   Dim trueVal As pdfValue
'   Set trueVal = pdfValue.TrueValue()
'   Debug.Print trueVal.ToString() ' Outputs: "true"
Public Function TrueValue() As pdfValue
    If Not Me Is pdfValue Then
        Set TrueValue = pdfValue.TrueValue
        Exit Function
    End If
    
    Static m_TrueInstance As pdfValue
    If m_TrueInstance Is Nothing Then
        Set m_TrueInstance = New pdfValue
        m_TrueInstance.valueType = PDF_Boolean
        m_TrueInstance.value = True
        'm_TrueInstance.m_id = 0
        'm_TrueInstance.m_generation = 0
    End If
    Set TrueValue = m_TrueInstance
End Function

' Returns a PDF boolean object with value False as defined in PDF specification 7.3.2.
'
' PDF boolean objects represent logical true/false values using keywords "true" and "false".
' Used in PDF dictionaries for flags, options, and conditional settings.
' Boolean values are case-sensitive and must be lowercase in PDF syntax.
' This method always returns the same False boolean instance, if used consistently
' then two true objects can be directly compared obj1 Is obj2.
'
' Returns:
'   pdfValue - A pdfValue instance with ValueType = PDF_Boolean and Value = False.
'
' Usage:
'   Dim falseVal As pdfValue
'   Set falseVal = pdfValue.FalseValue()
'   Debug.Print falseVal.ToString() ' Outputs: "false"
Public Function FalseValue() As pdfValue
    If Not Me Is pdfValue Then
        Set FalseValue = pdfValue.FalseValue
        Exit Function
    End If
    
    Static m_FalseInstance As pdfValue
    If m_FalseInstance Is Nothing Then
        Set m_FalseInstance = New pdfValue
        m_FalseInstance.valueType = PDF_Boolean
        m_FalseInstance.value = False
        'm_FalseInstance.m_id = 0
        'm_FalseInstance.m_generation = 0
    End If
    Set FalseValue = m_FalseInstance
End Function

' Creates a PDF value of the specified type.
'
' The valueType parameter determines which optional parameters are required or used.
' Supports all PDF data types as defined in PDF specification sections 7.3.2 through 7.3.10.
'
' Parameters:
'   valueType (PDF_ValueType) - The type of PDF value to create.
'    Range: Any valid PDF_ValueType enum value.
'    Expected: One of PDF_Null, PDF_Boolean, pdfNumber, PDF_String, PDF_Name, PDF_Array, PDF_Dictionary, PDF_Reference, PDF_Object.
'
'   value (Optional Variant) - The primary value content.
'    Required for: PDF_Boolean, PDF_Integer, PDF_Real, PDF_String, PDF_Name.  This is value of indicated type.
'    Required for: PDF_Reference.  This is the objID of object referenced.  Note: generation is set to 0.
'    Required for: PDF_Object. This should be another pdfValue that represents the object's content.
'    Ignored for: PDF_Null, PDF_Array, PDF_Dictionary.
'    Range: Boolean for PDF_Boolean, Numeric for pdfNumber, String for PDF_String/PDF_Name, Long for PDF_Reference, pdfValue for PDF_Object
'
'   objId (Optional Long) - Object identifier for PDF_Object type.
'    Range: 1 to 2,147,483,647 for valid object numbers.
'    Required only for PDF_Object type.
'
'   generation (Optional Long) - Generation number for PDF_Object type.
'    Range: 0 to 65,535, typically 0 for new objects.
'    Optional for PDF_Object type, defaults to 0.
'
' Returns:
'   pdfValue - A new pdfValue instance of the specified type, or shared instance for booleans.
'
' Usage Examples:
'   ' Create null value
'   Set nullVal = pdfValue.NewValue(PDF_Null)
'
'   ' Create boolean values
'   Set trueVal = pdfValue.NewValue(PDF_Boolean, True)
'   Set falseVal = pdfValue.NewValue(PDF_Boolean, False)
'
'   ' Create numeric values
'   Set intVal = pdfValue.NewValue(PDF_Integer, 42)
'   Set realVal = pdfValue.NewValue(PDF_Real, 3.14159)
'
'   ' Create string values
'   Set strVal = pdfValue.NewValue(PDF_String, "Hello World")
'
'   ' Create name values
'   Set nameVal = pdfValue.NewValue(PDF_Name, "Type")
'   Set nameVal = pdfValue.NewValue(PDF_Name, "/Type")
'
'   ' Create collection values
'   Set arrVal = pdfValue.NewValue(PDF_Array)
'   Set dictVal = pdfValue.NewValue(PDF_Dictionary)
'
'   ' Create reference values
'   Set refVal = pdfValue.NewValue(PDF_Reference, objId:=5)
'   Set refVal = pdfValue.NewValue(PDF_Reference, myObj,myObj.ID, myObj.generation)
'
'   ' Create object values
'   Dim contentDict As pdfValue
'   Set contentDict = pdfValue.NewValue(PDF_Dictionary)
'   Set objVal = pdfValue.NewValue(PDF_Object, contentDict, 5, 0)
Public Function NewValue(ByVal valueType As PDF_ValueType, _
    Optional ByVal value As Variant, _
    Optional ByVal objId As Long = 0, _
    Optional ByVal generation As Long = 0) As pdfValue
    
    Select Case valueType
        Case PDF_Null
            Set NewValue = NullValue()
        
        Case PDF_Boolean
            If IsMissing(value) Then
                Set NewValue = FalseValue()
            ElseIf CBool(value) Then
                Set NewValue = TrueValue()
            Else
                Set NewValue = FalseValue()
            End If
        
        Case PDF_ValueType.PDF_Integer, PDF_ValueType.PDF_Real, PDF_ValueType.PDF_String, PDF_ValueType.PDF_Name, PDF_ValueType.PDF_Comment
            If IsMissing(value) Then
                Err.Raise 5, "pdfValue.NewValue", "Value parameter required for " & typeName(valueType)
            End If
            
            Set NewValue = New pdfValue
            NewValue.valueType = valueType
            NewValue.id = objId ' 0
            NewValue.generation = generation ' 0
            
            Select Case valueType
                Case PDF_Integer
                    NewValue.value = CLng(value)
                Case PDF_Real
                    NewValue.value = CDbl(value)
                Case PDF_Name
                    ' Ensure name values start with forward slash
                    If Left(CStr(value), 1) <> "/" Then
                        NewValue.value = "/" & CStr(value)
                    Else
                        NewValue.value = CStr(value)
                    End If
                Case Else ' PDF_String, PDF_Comment
                    NewValue.value = value
            End Select
        
        Case PDF_Array
            Set NewValue = New pdfValue
            NewValue.valueType = PDF_Array
            ' if no array value supplied then create empty collection for array
            If typeName(value) = "Collection" Then
                Set NewValue.value = value
            Else
                Set NewValue.value = New Collection
                If typeName(value) = "pdfValue" Then
                    NewValue.value.Add value
                ElseIf Not IsMissing(value) Then
                    Err.Raise 5, "pdfValue.NewValue", "Value parameter must be Collection or pdfValue for PDF_Array!"
                End If
            End If
            NewValue.id = objId
            NewValue.generation = generation
        
        Case PDF_Dictionary
            Set NewValue = New pdfValue
            NewValue.valueType = PDF_Dictionary
            ' if no dictionary value supplied then create empty dictionary
            If typeName(value) = "Dictionary" Then
                Set NewValue.value = value
            Else
                Set NewValue.value = New Dictionary
            End If
            NewValue.id = objId
            NewValue.generation = generation
        
        Case PDF_Reference
            If objId = 0 Then
                Err.Raise 5, "pdfValue.NewValue", "Object ID required for reference type"
            End If
            
            Set NewValue = New pdfValue
            NewValue.valueType = PDF_Reference
            If IsMissing(value) Then
                Set NewValue.value = Nothing
            Else
                Set NewValue.value = value    ' should be pdfValue, the pdf object this is indirect reference for
            End If
            NewValue.id = objId
            NewValue.generation = generation
        
        Case PDF_Object
            If IsMissing(value) Then
                Err.Raise 5, "pdfValue.NewValue", "Content value required for object type"
            End If
            If objId = 0 Then
                Err.Raise 5, "pdfValue.NewValue", "Valid object ID required for object type"
            End If
            
            Set NewValue = New pdfValue
            NewValue.valueType = PDF_Object
            If IsObject(value) Then
                Set NewValue.value = value
            Else
                Let NewValue.value = value
            End If
            NewValue.id = objId
            NewValue.generation = generation
        
        Case Else
            Err.Raise 5, "pdfValue.NewValue", "Unsupported value type: " & valueType
    End Select
End Function

' Creates a PDF dictionary with specified /Type and optional /Subtype entries as defined in PDF specification.
'
' This convenience method simplifies creating typed PDF dictionaries that are fundamental to PDF structure.
' Most PDF objects are dictionaries with a /Type entry that identifies their purpose (Page, Catalog, Font, etc.).
' Many also have a /Subtype entry for further classification (Type1, TrueType fonts, etc.).
' The method automatically creates the dictionary, sets the /Type entry, and optionally sets /Subtype.
'
' Parameters:
'   typeName (String) - The PDF object type name for the /Type dictionary entry.
'    Range: Any valid PDF type name string.
'    Expected: Standard PDF types like "Page", "Catalog", "Font", "XObject", "Annot", etc.
'    Note: Leading slash is added automatically if not present.
'
'   subTypeName (Optional String) - The PDF object subtype name for the /Subtype dictionary entry.
'    Range: Any valid PDF subtype name string.
'    Expected: Subtype values like "Type1", "TrueType", "Image", "Form", "Text", etc.
'    Note: Leading slash is added automatically if not present. Omit to skip /Subtype entry.
'
' Returns:
'   pdfValue - A new pdfValue dictionary with ValueType = PDF_Dictionary containing /Type and optional /Subtype entries.
'
' Usage Examples:
'   ' Create a page dictionary
'   Dim pageDict As pdfValue
'   Set pageDict = pdfValue.NewTypedDictionary("Page")
'   ' Equivalent to: Set pageDict = pdfValue.NewValue(PDF_Dictionary)
'   '    pageDict.SetDictionaryValue "/Type", pdfValue.GetInternedName("Page")
'
'   ' Create a font dictionary with subtype
'   Dim fontDict As pdfValue
'   Set fontDict = pdfValue.NewTypedDictionary("Font", "Type1")
'   ' Equivalent to: Set fontDict = pdfValue.NewValue(PDF_Dictionary)
'   '    fontDict.SetDictionaryValue "/Type", pdfValue.GetInternedName("Font")
'   '    fontDict.SetDictionaryValue "/Subtype", pdfValue.GetInternedName("Type1")
'
'   ' Create an annotation dictionary
'   Dim annotDict As pdfValue
'   Set annotDict = pdfValue.NewTypedDictionary("Annot", "Text")
'
'   ' Create a catalog dictionary (no subtype needed)
'   Dim catalogDict As pdfValue
'   Set catalogDict = pdfValue.NewTypedDictionary("Catalog")
Public Function NewTypedDictionary(ByVal typeName As String, _
    Optional ByVal subTypeName As String = vbNullString) As pdfValue
    
    ' Create the base dictionary
    Dim dictVal As pdfValue
    Set dictVal = NewValue(PDF_Dictionary)
    
    ' Set the /Type entry using interned name for efficiency
    dictVal.SetDictionaryValue "/Type", GetInternedName(typeName)
    
    ' Set the /Subtype entry if provided
    If subTypeName <> vbNullString Then
        dictVal.SetDictionaryValue "/Subtype", GetInternedName(subTypeName)
    End If
    
    Set NewTypedDictionary = dictVal
End Function

' Returns an interned PDF name object to avoid creating duplicate objects and
' enable simplified comparisons via VBA Is keyword of instances.
'
' Interned names should be used for frequently used PDF names such as "/Type",
' "/Pages", "/Outline", "/Font", etc. The same name string always
' returns the same object instance, enabling fast reference equality checks.
'
' Parameters:
'   nameValue (String) - The name content, with or without leading slash.
'    Range: Any valid identifier string.
'    Expected: Standard PDF names like "Type", "Pages", "Outline".
'    Note: Leading slash is added automatically if not present.
'
' Returns:
'   pdfValue - A cached pdfValue instance with ValueType = PDF_Name.
'    Same instance returned for identical name strings.
'
' Usage:
'   Dim typeVal1 As pdfValue, typeVal2 As pdfValue
'   Set typeVal1 = pdfValue.GetInternedName("Type")
'   Set typeVal2 = pdfValue.GetInternedName("/Type")
'   Debug.Print (typeVal1 Is typeVal2) ' Outputs: True (same instance)
'
'   ' Common usage in PDF dictionaries
'   dict.SetDictionaryValue "/Type", pdfValue.GetInternedName("Page")
Public Function GetInternedName(ByVal nameValue As String) As pdfValue
    ' our cache that ensures unique names return same objects is in implicit instance
    If Not Me Is pdfValue Then
        Set GetInternedName = pdfValue.GetInternedName(nameValue)
        Exit Function
    End If
    
    ' Ensure name starts with forward slash
    If Left(nameValue, 1) <> "/" Then
        nameValue = "/" & nameValue
    End If
    
    ' on first call create our cache
    If m_InternedNames Is Nothing Then Set m_InternedNames = New Dictionary
    
    ' Check if already interned
    If m_InternedNames.Exists(nameValue) Then
        Set GetInternedName = m_InternedNames(nameValue)
    Else
        ' Create new interned name
        Dim nameVal As pdfValue
        Set nameVal = pdfValue.NewValue(PDF_Name, nameValue)
        
        ' Store in interned names dictionary
        Set m_InternedNames(nameValue) = nameVal
        Set GetInternedName = nameVal
    End If
End Function

' These Functions provide convenience routines to manipulate a pdfValue instance invoked from.

' Returns the count of elements in the PDF value based on its type.
'
' Provides consistent interface for determining the size or element count of PDF values.
' For arrays and dictionaries, returns the actual number of contained elements.
' For null values, returns 0 to indicate emptiness.
' For all other scalar types (boolean, integer, real, string, name, reference, object), returns 1.
' Note: arrays are stored as Collection objects so index is 1 based
'
' Returns:
'   Long - The count of elements:
'    0 for PDF_Null values
'    1 for scalar values (PDF_Boolean, pdfNumber, PDF_String, PDF_Name, PDF_Reference, PDF_Object)
'    Actual element count for PDF_Array and PDF_Dictionary
'    For PDF_Object: returns count of the contained value
'
' Usage:
'   Dim arr As pdfValue
'   Set arr = pdfValue.NewValue(PDF_Array)
'   arr.Add pdfValue.NewValue(pdfNumber, 1)
'   arr.Add pdfValue.NewValue(pdfNumber, 2)
'   Debug.Print arr.Count ' Outputs: 2
'
'   Dim dict As pdfValue
'   Set dict = pdfValue.NewValue(PDF_Dictionary)
'   dict.SetDictionaryValue "/Type", pdfValue.GetInternedName("Page")
'   Debug.Print dict.Count ' Outputs: 1
'
'   Debug.Print pdfValue.NullValue().Count ' Outputs: 0
'   Debug.Print pdfValue.NewValue(pdfNumber, 42).Count ' Outputs: 1
Public Property Get count() As Long
    Select Case valueType
        Case PDF_Null
            count = 0
        Case PDF_Array
            If typeName(value) = "Collection" Then
                count = value.count
            Else
                count = 0
            End If
        Case PDF_Dictionary
            If typeName(value) = "Dictionary" Then
                count = value.count
            Else
                count = 0
            End If
        Case PDF_Object
            ' For objects, return count of contained value
            If typeName(value) = "pdfValue" Then
                Dim innerValue As pdfValue
                Set innerValue = value
                count = innerValue.count
            Else
                count = 1
            End If
        Case Else
            count = 1
    End Select
End Property

' Converts the PDF value to its string representation for PDF serialization.
'
' Generates properly formatted PDF syntax for the value according to PDF specification.
' Output format depends on ValueType: null, true/false, numbers, (strings), /names,
' [arrays], <<dictionaries>>, "obj gen R" for references, or object definitions for objects.
' String output is suitable for direct inclusion in PDF content streams or object definitions
' but is primarily meant for debugging, see serialize for output to files.
'
' Returns:
'   String - The PDF-formatted representation of the value.
'    Empty string returned for unrecognized value types.
'
' Usage:
'   Debug.Print myValue.ToString()
'   fileContent = fileContent & myValue.ToString()
' returns vba String representation of value
' Warning: returns Nothing if object is not compatible with String type
Public Function ToString() As String
    On Error GoTo errHandler
    
    Select Case valueType
        Case PDF_ValueType.PDF_Null
            ToString = "null"
        Case PDF_ValueType.PDF_Boolean
            If CBool(value) Then
                ToString = "true"
            Else
                ToString = "false"
            End If
        Case PDF_ValueType.PDF_Integer
            ToString = Format(CLng(value), "0")
        Case PDF_ValueType.PDF_Real
            ToString = CDbl(value) ' don't format as we want all current digits stored, it won't add extra 0s anyway
            ' ensure has .0 if whole number
            If InStr(1, ToString, ".", vbBinaryCompare) < 1 Then ToString = ToString & ".0"
        Case PDF_ValueType.PDF_Reference
            ToString = m_id & " " & m_generation & " R"
        Case PDF_ValueType.PDF_Comment
            ToString = "%" & CStr(value) & vbLf
        Case Else ' PDF_Name, PDF_String, PDF_Array, PDF_Dictionary, PDF_Object, PDF_Operator, ...
            ToString = BytesToString(serialize())
    End Select
    
    Exit Function
errHandler:
    Debug.Print Err.Description & " (" & Err.Number & ")"
    Stop
End Function

' Escapes special characters in PDF strings
Private Function EscapeString(ByVal text As String) As String
    Dim result As String
    result = Replace(text, "\", "\\")
    result = Replace(result, "(", "\(")
    result = Replace(result, ")", "\)") ' (only unbalanced ) actually need escaping)
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, Chr(9), "\t")
    EscapeString = result
End Function

' Helper function to convert pdf string to byte array (0..N-1)
' escaping values as needed
Private Function PDF_StringToBytes(ByRef str As String, Optional ByVal asHex As Boolean = False) As Byte()
    On Error GoTo errHandler
    Dim strLen As Long
    Dim ucBytes() As Byte, bytes() As Byte
    Dim escapedStr As String
    escapedStr = EscapeString(str)
    strLen = Len(escapedStr)
    If strLen < 1 Then
        PDF_StringToBytes = StringToBytes("()") ' empty PDF string
        Exit Function
    End If
    
    ' we need to check if any Unicode values and if so we encode in UTF-8
    ' to avoid more expensive check if >255 value used, we simply always encode as UTF-8 unless ASCII (all values < 128)
    ucBytes = StringToUtf8Bytes(escapedStr, BOM:=False)  ' Note: if only ASCII then equivalent to UTF-8 byte array
    If UBound(ucBytes) > strLen Then ucBytes = addUtf8BOM(ucBytes)    ' something >128 and had to be expanded to multiple bytes
    
    Dim ndx As Long
    ' now determine if we are returning bytes as (string) or <## ## ##> format
    If asHex Then
        ' we need space for <## ##> assume 3 bytes per byte + 2 for enclosing <> - 1 as last one doesn't need space after
        ReDim bytes(0 To ((UBound(ucBytes) + 1) * 3)) ' Note: we check above for empty string, so can assume at least 1 byte
        bytes(0) = Asc("<")
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
        bytes(UBound(bytes)) = Asc(">") ' overwrites space from last "## " in loop
    Else
        ReDim bytes(0 To UBound(ucBytes) + 2)
        bytes(0) = Asc("(")
        bytes(UBound(bytes)) = Asc(")")
        CopyBytes ucBytes, bytes, 0, 1
    End If
    
    PDF_StringToBytes = bytes
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function

' a Name may have value encoded in #00 form, ie # followed by 2 hex digits
' from spec all Names should be UTF-8 encoded, so always convert from UTF-8 (BOM or not)
Public Function UnEscapeName(ByRef name() As Byte) As String
    Dim buffer() As Byte
    ReDim buffer(0 To UBound(name))
    Dim nameNdx As Long, bufNdx As Long, val As Long
    For nameNdx = LBound(name) To UBound(name)
        val = name(nameNdx)
        If val = 35 Then ' #
            Dim s As String
            s = "&H" & Chr(name(nameNdx + 1)) & Chr(name(nameNdx + 2))  ' #00 --> &H00
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
        UnEscapeName = "/" & Mid(Utf8BytesToString(buffer), 3) ' skip past BOM [/ BOM Name]
    Else
        UnEscapeName = Utf8BytesToString(buffer)
    End If
End Function

' opposite of UnEscapeName, converts spaces and other weird characters into #00 form
' from spec all Names should be UTF-8 encoded, so no need for a BOM
Public Sub EscapeName(ByRef name As String, ByRef destBuffer() As Byte)
    ' copy over each character, but ensuring control codes <=32 & non ASCII >= 128 are encoded as hex digits
    ' we assume a String in VBA is a Unicode String so convert name to UTF8 bytes
    Dim utf8Name() As Byte: utf8Name = StringToUtf8Bytes(name)
    Dim srcNdx As Long, dstNdx As Long
    Dim b As Byte
    Dim destLen As Long
    ' add room for leading / if not present
    If utf8Name(0) <> 47 Then destLen = 1 ' else destlen = 0 ' 47="/"
    ' calc our destination buffer size
    For srcNdx = LBound(utf8Name) To UBound(utf8Name)
        b = utf8Name(srcNdx)
        If (b <= 32) Or (b >= 127) Then   ' from spec < Asc("!") or > Asc("~") need encoding
            destLen = destLen + 3 ' + #00, the hex encoded value
        Else
            destLen = destLen + 1 ' + the character
        End If
    Next srcNdx
    ' allocate destination buffer and copy over escaped
    ReDim destBuffer(0 To (destLen - 1))
    ' ensure buffer starts with /
    destBuffer(0) = 47
    dstNdx = 1 ' skip past / we stored
    If utf8Name(0) <> 47 Then
        srcNdx = LBound(utf8Name)
    Else
        srcNdx = LBound(utf8Name) + 1 ' we already have starting / so skip if source also starts with it
    End If
    For srcNdx = srcNdx To UBound(utf8Name)
        b = utf8Name(srcNdx)
        If (b <= 32) Or (b >= 127) Then   ' from spec < Asc("!") or > Asc("~") need encoding
            destBuffer(dstNdx) = 35 ' #
            destBuffer(dstNdx + 1) = Hex(b \ 16)   ' high nibble
            destBuffer(dstNdx + 2) = Hex(b Mod 16) ' low nibble
            dstNdx = dstNdx + 3
        Else
            destBuffer(dstNdx) = b
            dstNdx = dstNdx + 1
        End If
    Next srcNdx
    ' cleanup and return our escaped name
    Erase utf8Name
    'EscapeName = Utf8BytesToString(destBuffer)
    'Erase destBuffer
End Sub

' convert pdf Value (Me) into a Byte() array as stored in pdf file
Public Function serialize() As Byte()
    On Error GoTo errHandler
    Dim objBytes() As Byte
    Dim v As Variant
    Dim pv As pdfValue
    Dim firstPass As Boolean
    Dim chunks() As Variant
    Dim ndx As Long, bigChunkSize As Long
    Dim dstOffset As Long
    Dim str As String
    
    Select Case valueType
        Case PDF_ValueType.PDF_Name
            str = CStr(value)
            EscapeName str, objBytes  ' see below where Dictionary Key Name encoded
        Case PDF_ValueType.PDF_String
            str = CStr(value)
            ' convert to bytes and escape values as needed
            objBytes = PDF_StringToBytes(str, (flags And ValueFlags.flgBinary) = ValueFlags.flgBinary)
        Case PDF_ValueType.PDF_Array
            ' our array bracketing
            objBytes = StringToBytes("[ ")
            ' convert each array element to a Byte() array
            For Each v In value
                Set pv = v
                CopyBytes pv.serialize(), objBytes, 0, UBound(objBytes) + 1
                CopyBytes StringToBytes(" "), objBytes, 0, UBound(objBytes) + 1
            Next v
            CopyBytes StringToBytes("]"), objBytes, 0, UBound(objBytes) + 1
        Case PDF_ValueType.PDF_Dictionary
            objBytes = StringToBytes("<<" & vbLf)
            Dim dict As Dictionary
            Set dict = value
            For Each v In dict.keys
                Dim key As String
                If typeName(v) = "String" Then
                    key = CStr(v)
                Else 'If typeName(v) = "pdfValue" Then ' assume pdfValue with valueType=PDF_Name
                        key = CStr(v.value)
                End If
                ' key is always a /Name and should be considered UTF8 without BOM
                Dim keyBuffer() As Byte
                EscapeName key, keyBuffer
                Set pv = dict.item(v) ' v may be pdfValue or String version of key
                Dim dataBuffer() As Byte
                dataBuffer = pv.serialize()
                ndx = ByteArraySize(objBytes) ' get current size, offset to start copy to
                ReDim Preserve objBytes(0 To (ndx + ByteArraySize(keyBuffer) + ByteArraySize(dataBuffer) + 1)) ' +2 for space & linefeed -1 array bound
                CopyBytes keyBuffer, objBytes, 0, ndx
                ndx = ndx + ByteArraySize(keyBuffer)
                objBytes(ndx) = 32 ' space
                CopyBytes dataBuffer, objBytes, 0, ndx + 1
                ndx = ndx + ByteArraySize(dataBuffer) + 1
                objBytes(ndx) = 10 ' vbLf
                Erase keyBuffer
                Erase dataBuffer
            Next v
            ndx = ByteArraySize(objBytes)
            ReDim Preserve objBytes(0 To (ndx + 2)) ' +3 for >>\n -1 for array bound
            CopyBytes StringToBytes(">>" & vbLf), objBytes, 0, ndx
        Case PDF_ValueType.PDF_Stream       ' actual stream object with dictionary and data
            Dim stream As pdfStream
            Set stream = value
            objBytes = stream.stream_meta.serialize()
            CopyBytes stream.stream_data.serialize(), objBytes, 0, UBound(objBytes) + 1
        Case PDF_ValueType.PDF_StreamData   ' represents only stream ... endstream portion
            objBytes = StringToBytes("stream" & vbLf)
            CopyBytes Me.value, objBytes, 0, UBound(objBytes) + 1
            CopyBytes StringToBytes(vbLf & "endstream" & vbLf), objBytes, 0, UBound(objBytes) + 1
        
    ' to simplify processing, not one of 9 basic types either
        Case PDF_ValueType.PDF_Object       ' id generation obj << dictionary >> endobj
            objBytes = StringToBytes(m_id & " " & m_generation & " obj" & vbLf)
            Set pv = value
            CopyBytes pv.serialize(), objBytes, 0, UBound(objBytes) + 1
            'If objBytes(UBound(objBytes)) <> 10 Then CopyBytes StringToBytes(vbLf), objBytes, 0, UBound(objBytes) + 1
            CopyBytes StringToBytes("endobj" & vbLf), objBytes, 0, UBound(objBytes) + 1
        Case PDF_ValueType.PDF_Trailer
            Set pv = value
            objBytes = StringToBytes("trailer" & vbLf)
            CopyBytes pv.serialize(), objBytes, 0, UBound(objBytes) + 1
        Case PDF_ValueType.PDF_Operator
            objBytes = StringToBytes(CStr(value))
        Case PDF_ValueType.PDF_EndOfArray, PDF_ValueType.PDF_EndOfDictionary, PDF_ValueType.PDF_EndOfObject, PDF_ValueType.PDF_EndOfStream
            ' return Nothing, not actual values but markers
        Case PDF_ValueType.PDF_Null, PDF_ValueType.PDF_Boolean, PDF_ValueType.PDF_Integer, PDF_ValueType.PDF_Real, PDF_ValueType.PDF_Reference, PDF_ValueType.PDF_Comment
            objBytes = StringToBytes(ToString())
        Case Else
            Debug.Print "Warning: unexpected PDF_ValueType of " & valueType & ", returning """""
            Stop ' unexpected PDF_Value_Type
    End Select
    
    serialize = objBytes
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function



' Adds one or more items to the end of a PDF array value using VBA Collection.Add method.
'
' Appends the specified item(s) to the end of the array if this value is of type PDF_Array.
' The item(s) can be any PDF value type including other arrays and dictionaries.
' Operation is ignored if this value is not an array type.
' Array maintains insertion order and allows duplicate values.
' Uses VBA Collection.Add method which appends to the end of the collection.
' Note: Arrays are stored in VBA Collection objects so use 1-based indexing
' Note: if item(s) are not pdfValue objects, then automatically converted to pdfValue
'
' Parameters:
'   items (ParamArray Variant) - Variable number of items to add to the array.
'    Range: Any combination of numeric, string, boolean, or pdfValue types.
'    Expected: Typically pdfValue objects for proper PDF structure.
'    Note: Items are added in the order specified.
'
' Usage:
'   Dim arr As pdfValue
'   Set arr = pdfValue.NewValue(PDF_Array)
'   arr.Add pdfValue.NewValue(pdfNumber, 42)
'   arr.Add pdfValue.NewValue(PDF_String, "text")
'   arr.Add pdfValue.GetInternedName("Type")
'   Debug.Print arr.Count ' Outputs: 3
'
'   Dim arr As pdfValue
'   Set arr = pdfValue.NewValue(pdfArray)
'   ' Add multiple items at once
'   arr.AddItems 1, 2, 3, "text", True, "/Name"
'   ' Equivalent to multiple individual Add calls:
'   ' arr.Add pdfValue.NewValue(pdfInteger, 1)
'   ' arr.Add pdfValue.NewValue(pdfInteger, 2)
'   ' arr.Add pdfValue.NewValue(pdfInteger, 3)
'   ' arr.Add pdfValue.NewValue(pdfString, "text")
'   ' arr.Add pdfValue.TrueValue()
'   ' arr.Add pdfValue.GetInternedName("/Name")
'
'   ' Build array incrementally
'   Dim coords As pdfValue
'   Set coords = pdfValue.NewValue(pdfArray)
'   coords.AddItems 0, 0        ' Add first two coordinates
'   coords.AddItems 612, 792    ' Add remaining coordinates
Public Sub Add(ParamArray items() As Variant)
    ' we want to allow calling directly on PDF_Array or wrapping PDF_Object (or its indirect reference)
    ' however, if we recursively call value.Add items then items will be wrapped in another array
    ' so to avoid the nesting we determine the PDF_Array object which might be Me or a child and work with it
    Dim pdfArrObj As pdfValue: Set pdfArrObj = Me

    ' if called on object or indirect object then pass call on its array value
    If valueType = PDF_ValueType.PDF_Object Then
        Set pdfArrObj = value
    ElseIf valueType = PDF_ValueType.PDF_Reference Then
        If typeName(value) = "pdfValue" Then    ' if Not value Is Nothing and its a pdfValue
            Set pdfArrObj = value.value
        End If
    End If
    
    ' validate we have a PDF array
    If pdfArrObj Is Nothing Then
        Err.Raise 13, "pdfValue.Add", "Add called on Nothing, can only be called on PDF_Array values"
        Exit Sub
    End If
    
    ' the rest of this operates on our PDF_Array instance, which might be Me or a child of Me or referred to by Me
    With pdfArrObj
        If .valueType <> PDF_Array Or typeName(.value) <> "Collection" Then
            Err.Raise 13, "pdfValue.Add", "Add can only be called on PDF_Array values"
            Exit Sub
        End If
    
        ' Process each parameter
        Dim ndx As Long
        For ndx = LBound(items) To UBound(items)
            Dim currentItem As Variant
            If IsObject(items(ndx)) Then
                Set currentItem = items(ndx)
            Else
                Let currentItem = items(ndx)
            End If
        
            Dim pdfItem As pdfValue
            Set pdfItem = ConvertToPdfValue(currentItem)
        
            .value.Add pdfItem
        Next ndx
    End With
End Sub




' Gets a value from a PDF array by 1-based index.
'
' Retrieves the value at the specified 1-based index from the array.
' Returns the actual value if index is valid, otherwise returns Nothing.
' Properly handles both object and non-object return values.
' Operation returns Nothing if this value is not an array type.
' Note: Arrays use 1-based indexing as represented with VBA Collection objects.
'
' Parameters:
'   index (Long) - The 1-based index of the array element to retrieve.
'    Range: 1 to Count for valid access.
'    Expected: Positive integer within array bounds.
'
' Returns:
'   Variant - The value at the specified index, or Nothing if index invalid or not an array.
'    May be pdfValue object or other variant types.
'
' Usage:
'   Dim arr As pdfValue
'   Set arr = pdfValue.NewValue(PDF_Array)
'   arr.Add pdfValue.NewValue(pdfNumber, 42)
'   arr.Add pdfValue.NewValue(PDF_String, "text")
'
'   Dim firstItem As pdfValue
'   Set firstItem = arr.GetArrayValue(1) ' Gets first element (1-based)
'   Debug.Print firstItem.ToString() ' Outputs: "42"
'
'   Dim secondItem As pdfValue
'   Set secondItem = arr.GetArrayValue(2) ' Gets second element
'   Debug.Print secondItem.ToString() ' Outputs: "(text)"
Public Function GetArrayValue(ByVal index As Long) As Variant
    Set GetArrayValue = Nothing
    
    ' if we are a pdf object with an array as a value, access nested array instance
    Dim pdfArrObj As pdfValue: Set pdfArrObj = Me
    If valueType = PDF_ValueType.PDF_Object Then
        Set pdfArrObj = value
    ElseIf valueType = PDF_ValueType.PDF_Reference Then
        If typeName(value) = "pdfValue" Then    ' if Not value Is Nothing and its a pdfValue
            Set pdfArrObj = value.value
        End If
    End If
    
    With pdfArrObj
        If .valueType = PDF_ValueType.PDF_Array And typeName(.value) = "Collection" Then
            If index >= 1 And index <= .value.count Then
                If IsObject(.value(index)) Then
                    Set GetArrayValue = .value(index)
                Else
                    Let GetArrayValue = .value(index)
                End If
            End If
        End If
    End With
End Function

' Creates a PDF array with variable number of items using ParamArray allowing for flexible initialization.
'
' Convenience metho to create PDF arrays with multiple values in a single call.
' Accepts any number of parameters of mixed types and automatically
' converts them to appropriate PDF values. Example uses include creating coordinate
' arrays, color arrays, and other multi-value PDF structures.
'
' Allow the raw value to be passed and automatically creates correct pdfValue wrapper:
' - Numeric values become PDF_Integer or PDF_Real based on decimal content
' - String values become PDF_String (literal strings) or PDF_Name (if prefixed with "/")
' - Boolean values become PDF_Boolean
' - Existing pdfValue objects are added directly
' - Other objects trigger an error
'
' Parameters:
'   items (ParamArray Variant) - Variable number of items to add to the array.
'    Range: Any combination of numeric, string, boolean, or pdfValue types.
'    Expected: Mixed types representing PDF array contents.
'    Note: Strings starting with "/" are treated as PDF names, others as literal strings.
'
' Returns:
'   pdfValue - A new pdfValue with ValueType = pdfArray containing all specified items.
'
' Usage Examples:
'   ' Create coordinate array for PDF rectangle [0 0 612 792]
'   Dim mediaBox As pdfValue
'   Set mediaBox = pdfValue.NewArray(0, 0, 612, 792)
'
'   ' Create mixed content array [3.4 0 0 52 (some string) (another string) 4.2 false true]
'   Dim mixedArray As pdfValue
'   Set mixedArray = pdfValue.NewArray(3.4, 0, 0, 52, "some string", "another string", 4.2, False, True)
'
'   ' Create color array for RGB [1.0 0.5 0.0]
'   Dim rgbColor As pdfValue
'   Set rgbColor = pdfValue.NewArray(1.0, 0.5, 0.0)
'
'   ' Create array with PDF names [/Type /Page /Parent 3 0 R]
'   Dim pageRef As pdfValue
'   Set pageRef = pdfValue.NewValue(pdfReference, , 3, 0)
'   Dim pageArray As pdfValue
'   Set pageArray = pdfValue.NewArray("/Type", "/Page", "/Parent", pageRef)
'
'   ' Create transformation matrix [1 0 0 1 100 200]
'   Dim transform As pdfValue
'   Set transform = pdfValue.NewArray(1, 0, 0, 1, 100, 200)
'
'   ' Empty array creation
'   Dim emptyArray As pdfValue
'   Set emptyArray = pdfValue.NewArray()
'
'   ' Create from VBA Array() function
'   Dim coords As pdfValue
'   Set coords = pdfValue.NewArray(Array(0, 0, 612, 792))
'
'   ' Create from computed array
'   Dim values As Variant
'   values = Array(1.0, 2.0, 3.0, 4.0)
'   Dim numArray As pdfValue
'   Set numArray = pdfValue.NewArray(values)
'
'   ' Create from worksheet range values
'   Dim rangeValues As Variant
'   rangeValues = Range("A1:D1").Value  ' Returns 2D array
'   Dim rowArray As pdfValue
'   Set rowArray = pdfValue.NewArray(rangeValues)
Public Function NewArray(ParamArray items() As Variant) As pdfValue
    ' Create the base array
    Dim arrayVal As pdfValue
    Set arrayVal = NewValue(PDF_Array)
        
    ' Allow alternative of first argument itself being an array
    ' i.e. VBA Array(...) or Range(...) value
    If IsArray(items(LBound(items))) Then
        If LBound(items) <> UBound(items) Then
            Err.Raise 5, "pdfValue.NewArray", "Mixed Array() and non-Array() arguments not supported, only single Array argument supported!"
        End If
        items = items(LBound(items))
    End If
        
    ' Process each parameter
    Dim ndx As Long
    For ndx = LBound(items) To UBound(items)
        Dim currentItem As Variant
        If IsObject(items(ndx)) Then
            Set currentItem = items(ndx)
        Else
            Let currentItem = items(ndx)
        End If
        
        Dim pdfItem As pdfValue
        Set pdfItem = ConvertToPdfValue(currentItem)
        
        arrayVal.Add pdfItem
    Next ndx
    
    Set NewArray = arrayVal
End Function


' Internal helper method to convert various VBA types to appropriate pdfValue objects.
'
' Handles automatic detection and conversion of common VBA types to their
' corresponding PDF value representations.
'
' Type Conversion Rules:
' - Integers and Longs become PDF_Integer
' - Singles and Doubles become PDF_Real
' - Strings starting with "/" become PDF_Name (with slash preserved)
' - Other strings become PDF_String
' - Boolean values become PDF_Boolean (using shared True/False instances)
' - Existing pdfValue objects pass through unchanged
' - Other object types raise an error
'
' Parameters:
'   item (Variant) - The value to convert to a pdfValue object.
'    Range: Any VBA-compatible type.
'    Expected: Numeric, string, boolean, or existing pdfValue.
'
' Returns:
'   pdfValue - A pdfValue object representing the input value.
'
' Usage:
'   ' Internal use
'   Dim converted As pdfValue
'   Set converted = ConvertToPdfValue(42)        ' Returns pdfInteger
'   Set converted = ConvertToPdfValue(3.14)      ' Returns pdfReal
'   Set converted = ConvertToPdfValue("text")    ' Returns pdfString
'   Set converted = ConvertToPdfValue("/Name")   ' Returns pdfName
'   Set converted = ConvertToPdfValue(True)      ' Returns pdfBoolean
Function ConvertToPdfValue(ByRef item As Variant) As pdfValue
    Select Case VarType(item)
        Case vbInteger, vbLong
            Set ConvertToPdfValue = NewValue(PDF_Integer, item)
            
        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            Set ConvertToPdfValue = NewValue(PDF_Real, item)
            
        Case vbString
            Dim strValue As String
            strValue = CStr(item)
            
            ' Check if it's a PDF name (starts with "/")
            If Left(strValue, 1) = "/" Then
                Set ConvertToPdfValue = GetInternedName(strValue)
            Else
                Set ConvertToPdfValue = NewValue(PDF_String, strValue)
            End If
            
        Case vbBoolean
            Set ConvertToPdfValue = NewValue(PDF_Boolean, item)
            
        Case vbObject
            ' Check if it's already a pdfValue
            If typeName(item) = "pdfValue" Then
                Set ConvertToPdfValue = item
            ElseIf typeName(item) = "Dictionary" Then
                Set ConvertToPdfValue = NewValue(PDF_Dictionary, item)
            ElseIf typeName(item) = "Collection" Then
                Set ConvertToPdfValue = NewValue(PDF_Array, item)
            ElseIf typeName(item) = "Nothing" Then
                Set ConvertToPdfValue = NullValue()
            Else
                Err.Raise 13, "pdfValue.ConvertToPdfValue", _
                    "Unsupported object type: " & typeName(item) & ". Only pdfValue objects are supported."
            End If
            
        Case vbNull, vbEmpty
            Set ConvertToPdfValue = NullValue()
            
        Case Else
            Err.Raise 13, "pdfValue.ConvertToPdfValue", _
                "Unsupported data type: " & typeName(item) & " (VarType: " & VarType(item) & ")"
    End Select
End Function

' convenience function to perform a deep copy of a pdfValue
' This is useful when duplicating objects for mutation or to create a new PDF and update separate from original.
' Note: once copy is performed, any changes to returned object tree are not replicated to original
' however, interned objects are still shared and changes to one will effect all other copies
Public Function Clone(Optional ByRef objRefCache As Dictionary = Nothing) As pdfValue
    Dim newVal As Variant
    
    ' determine if object copy already exists and if so reuse it
    Dim objKey As String, refKey As String
    If (valueType = PDF_ValueType.PDF_Object Or valueType = PDF_ValueType.PDF_Reference) And (Not objRefCache Is Nothing) Then
        objKey = id & " " & generation
        refKey = objKey & " R"
        
        ' try with obj first
        If objRefCache.Exists(objKey) Then
            Set Clone = objRefCache(objKey)
            If valueType = PDF_ValueType.PDF_Reference Then
                Set Clone = Clone.referenceObj
                If Not objRefCache.Exists(refKey) Then objRefCache.Add refKey, Clone
            End If
            Exit Function
        End If
        
        ' now check for just reference
        If objRefCache.Exists(refKey) Then
            Set Clone = objRefCache(refKey)
            If valueType = PDF_ValueType.PDF_Reference Then
                Exit Function
            End If
            ' if reference has link to obj, we can use it otherwise need to go on and clone
            If typeName(Clone.value) = "pdfValue" Then
                Set Clone = Clone.value
                objRefCache.Add objKey, Clone ' above check already determine doesn't .Exists(objKey)
                Exit Function
            End If
        End If
    End If
    
    ' create a copy of our value if its an object, else just use if scaler
    If IsObject(value) Then
        Dim v As Variant
        Select Case typeName(value)
            Case "pdfValue"
                Set newVal = value.Clone(objRefCache)
            Case "Collection"
                Dim coll As Collection: Set coll = New Collection
                Set newVal = coll
                For Each v In value
                    coll.Add v.Clone(objRefCache)
                Next v
            Case "Dictionary"
                Dim dict As Dictionary: Set dict = New Dictionary
                Set newVal = dict
                For Each v In value.keys
                    dict.Add v, value.item(v).Clone(objRefCache)
                Next v
        End Select
    Else
        Let newVal = value
    End If
    
    ' create the copy, set value to a copy or asis
    Set Clone = pdfValue.NewValue(valueType, newVal, id, generation)
    
    ' update our cache
    If (Not objRefCache Is Nothing) And (id <> 0) Then
        objKey = id & " " & generation
        refKey = objKey & " R"
        
        If valueType = PDF_ValueType.PDF_Object Then
            If Not objRefCache.Exists(objKey) Then Set objRefCache(objKey) = Clone
            If objRefCache.Exists(refKey) Then
                Set Clone.referenceObj = objRefCache(refKey)
            Else
                Set objRefCache(refKey) = Clone.referenceObj
            End If
            Set Clone.referenceObj.value = Clone
        ElseIf valueType = PDF_ValueType.PDF_Reference Then
            If objRefCache.Exists(objKey) Then
                ' reuse existing object's reference object if it exists, but ensure it point back to actual object
                Set Clone = objRefCache(objKey).referenceObj
                Set Clone.value = objRefCache(objKey)
            End If
            If Not objRefCache.Exists(refKey) Then Set objRefCache(refKey) = Clone
        End If
    End If
End Function


' *** reading in values from PDF byte array

' given raw PDF file contents as Byte array and offset in the array to peak, returns value type at offset
' contentStream: enables check for known Operators in a content stream and returning PDF_Operator when found
' Note: while always calling with contentStream will work, it will add unnecessary delays due to the extra string matching used
Private Function GetValueType(ByRef bytes() As Byte, ByVal offset As Long, Optional contentStream As Boolean = False) As PDF_ValueType
    On Error GoTo errHandler
    GetValueType = PDF_ValueType.PDF_Null
    
    offset = SkipWhitespace(bytes, offset)
    If offset > UBound(bytes) Then Exit Function    ' return null type if end of data
    Dim token As String: token = Chr(bytes(offset))
    
    Dim tmpStr As String
    
    ' only look for operator matches if processing contentStream data
    If contentStream Then
        Dim op_offset As Long: op_offset = offset ' use a copy of offset so if not operator offset is unchanged
        Dim op As PdfTextOp
        tmpStr = GetWord(bytes, op_offset)
        If IsKnownOperator(tmpStr, op) Then
            GetValueType = PDF_ValueType.PDF_Operator + LShift(op, 8) ' low byte flags as op, high byte of word = op
            Exit Function
        End If
    End If
    
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
            offset = SkipWhitespace(bytes, offset)
            words(1) = GetWord(bytes, offset)
            offset = SkipWhitespace(bytes, offset)
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
            'Stop ' error! unexpected value
    End Select
    
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function

' Returns a value loaded for a PDF
' Updates offset to next non-whitespace byte after this value is loaded
' bytes() may be content() or uncompressed stream object stream buffer
' offset is within bytes() provided
' Note: meta is only used for stream object
' contentStream: enables check for known Operators in a content stream and returning PDF_Operator when found
' objRefCache is used to ensure object and reference object instances are not duplicated
' however, objRefCache is per pdfDocument and only apply to object and indirect object references
Function GetValue(ByRef bytes() As Byte, ByRef offset As Long, _
    Optional ByRef Meta As Dictionary = Nothing, _
    Optional ByVal contentStream As Boolean = False, _
    Optional ByRef objRefCache As Dictionary = Nothing) As pdfValue
    On Error GoTo errHandler
    DoEvents
    
    ' Skip whitespace, updating offset
    offset = SkipWhitespace(bytes, offset)
    
    ' exit early if no value to return possible
    If offset > UBound(bytes) Then
        Set GetValue = pdfValue.NullValue
        Exit Function
    End If
    
    ' Get value type once and cache
    Dim result As pdfValue
    Dim valueType As PDF_ValueType
    valueType = GetValueType(bytes, offset, contentStream)
    
    ' Cache first byte for multiple comparisons
    Dim firstByte As Byte: firstByte = bytes(offset)
    
    ' misc variables used in parsing/extracting
    Dim tmpStr As String
    Dim val As pdfValue

    ' only need to check low byte, high byte of word (&HFF00) stores operator type
    Select Case (valueType And &HFF)
        Case PDF_ValueType.PDF_Null
            ' Skip past "null"
            offset = offset + 4
            Set result = pdfValue.NullValue
            
        Case PDF_ValueType.PDF_Operator
            Set result = New pdfValue
            result.valueType = valueType
            result.value = GetWord(bytes, offset)
            
        Case PDF_ValueType.PDF_Name
            Dim count As Long
            Dim word() As Byte
            GetAsBytes bytes, offset, word, count
            Set result = GetInternedName(UnEscapeName(word))
            
        Case PDF_ValueType.PDF_Boolean
            ' we assume if first character t/f indicates value
            If firstByte = 116 Then ' "t" for "true"
                Set result = pdfValue.TrueValue
                offset = offset + 4 ' Skip "true"
            Else ' "f" for "false"
                Set result = pdfValue.FalseValue
                offset = offset + 5 ' Skip "false"
            End If
            
        Case PDF_ValueType.PDF_Integer, PDF_ValueType.PDF_Real
            tmpStr = GetWord(bytes, offset)
            Set result = pdfValue.NewValue(valueType, tmpStr)
            
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
            
            ' Pre-allocate string buffer to reasonable size
            Dim strBuffer() As Byte
            Dim strLen As Long
            ReDim strBuffer(0 To 1023) ' Start with 1KB buffer
            strLen = 0
            
            If firstByte = 40 Then ' "(" - literal string
                offset = offset + 1
                
                ' only need to escape unbalanced ), so need to keep track of balanced ()
                ' Note: could see ((\))) or (\(\)) - from real pdfs, \( does not start new matching set
                Dim openParens As Integer: openParens = 0
                Do While (openParens > 0) Or (bytes(offset) <> 41) ' Asc(")") - if not open (, then end of string found
                    ' Resize buffer if needed (double size)
                    If strLen > UBound(strBuffer) Then
                        ReDim Preserve strBuffer(0 To (UBound(strBuffer) + 1) * 2 - 1)
                    End If
                    
                    ' if ( or ) then update our tracking, but otherwise treat as normal character in string
                    If bytes(offset) = 40 Then ' "("
                        openParens = openParens + 1
                    ElseIf bytes(offset) = 41 Then ' ")"
                        openParens = openParens - 1
                    End If
                    
                    If bytes(offset) = 92 Then ' Asc("\") then escaped value
                        offset = offset + 1
                        Select Case bytes(offset)
                            Case 110 ' Asc("n") - \n
                                strBuffer(strLen) = 10 ' Asc(vbLf)
                                strLen = strLen + 1
                            ' \ immediately followed by newline treated as line continuation & ignored
                            ' we treat \<CR><NL>, \<CR> and \<NL> the same (note \<NL><CR> treated as <CR> )
                            Case 13 ' Asc(vbCr) - \r or \r\n
                                ' line continuation, ignore line break, <CR> or <CR><NL>
                                If bytes(offset + 1) = 10 Then offset = offset + 1
                            Case 10 ' Asc(vbLf)
                                ' line continuation, ignore line break
                            Case 114 ' Asc("r") - \r
                                strBuffer(strLen) = vbCr
                                strLen = strLen + 1
                            Case 116 ' Asc("t") - \t
                                strBuffer(strLen) = vbTab
                                strLen = strLen + 1
                            Case 102 ' Asc("f") - \f
                                strBuffer(strLen) = &HC  ' formfeed
                                strLen = strLen + 1
                            Case 98  ' Asc("b") - \b
                                strBuffer(strLen) = 8    ' backspace
                                strLen = strLen + 1
                            Case 92, 40, 41 ' Asc("\") or Asc("(") or Asc(")") - \\ \( \)
                                strBuffer(strLen) = bytes(offset) ' \()
                                strLen = strLen + 1
                            
                            Case 48 To 55 ' Octal digits 0-7
                                Dim octVal As Long: octVal = bytes(offset) - 48
                                Dim octCount As Integer: octCount = 1
                                If (offset + 1) <= UBound(bytes) Then
                                    Do While octCount < 3 And bytes(offset + 1) >= 48 And bytes(offset + 1) <= 55
                                        offset = offset + 1
                                        If (offset + 1) > UBound(bytes) Then Exit Do
                                        octVal = (octVal * 8) + (bytes(offset) - 48)
                                        octCount = octCount + 1
                                    Loop
                                End If
                                strBuffer(strLen) = octVal
                                strLen = strLen + 1
                            Case Else
                                ' Unknown escape - treat as literal
                                strBuffer(strLen) = bytes(offset)
                                strLen = strLen + 1
                        End Select
                    Else
                        strBuffer(strLen) = bytes(offset)
                        strLen = strLen + 1
                    End If
                    DoEvents
                    offset = offset + 1
                Loop
                offset = offset + 1 ' skip past closing ")"
            
            Else  ' < hex encoded string >
                offset = offset + 1
                Do While bytes(offset) <> 62 ' Asc(">")
                    ' get 2 hex digits, ignoring whitespace, may end with odd # of hex digits
                    offset = SkipWhitespace(bytes, offset)
                    ' ensure enough space for our string, doubling size every time we exceed current allocation
                    If strLen > UBound(strBuffer) Then
                        ReDim Preserve strBuffer(0 To (UBound(strBuffer) + 1) * 2 - 1)
                    End If
                    
                    Dim hexVal As Integer
                    Dim hex1 As Byte: hex1 = bytes(offset)
                    offset = SkipWhitespace(bytes, offset + 1)
                    hexVal = HexDigitValue(hex1) * 16 ' high nibble
                    
                    If bytes(offset) <> 62 Then ' Not end of hex string
                        Dim hex2 As Byte: hex2 = bytes(offset)
                        offset = offset + 1
                        hexVal = hexVal + HexDigitValue(hex2)
                    'Else
                    '   hexVal = hexVal + 0 ' Pad with 0
                    End If
                    
                    strBuffer(strLen) = hexVal
                    strLen = strLen + 1
                    
                    DoEvents
                Loop
                offset = offset + 1 ' Skip past closing ">"
            End If
            
            ' we don't resize strBuffer, instead limiting conversion of valid bytes to string via strLen
            Set result = pdfValue.NewValue(PDF_ValueType.PDF_String, ProcessStringBuffer(strBuffer, strLen))
            
            ' mark if originally read in a hex <##> represenation or standard (...) form
            If firstByte = 40 Then result.flags = result.flags Or ValueFlags.flgBinary
            
        Case PDF_ValueType.PDF_Array
            offset = offset + 1 ' skip the opening [
            Dim col As Collection
            Set col = New Collection
            Set result = pdfValue.NewValue(valueType, col)
            Set val = GetValue(bytes, offset, contentStream:=contentStream, objRefCache:=objRefCache)
            Do While val.valueType <> PDF_ValueType.PDF_EndOfArray
                col.Add val
                Set val = pdfValue.GetValue(bytes, offset, contentStream:=contentStream, objRefCache:=objRefCache)
            Loop
            Set col = Nothing
        
        Case PDF_ValueType.PDF_Dictionary
            offset = offset + 2 ' skip past <<
            Dim dict As Dictionary
            Set dict = New Dictionary
            Set result = pdfValue.NewValue(valueType, dict)
            Dim name As pdfValue
            Set name = pdfValue.GetValue(bytes, offset, contentStream:=contentStream, objRefCache:=objRefCache)
            Do While name.valueType = PDF_ValueType.PDF_Name
                Set val = pdfValue.GetValue(bytes, offset, contentStream:=contentStream, objRefCache:=objRefCache)
                dict.Add CStr(name.value), val
                ' get name of next element of dictionary (or end of dictionary marker >> )
                Set name = pdfValue.GetValue(bytes, offset, contentStream:=contentStream, objRefCache:=objRefCache)
            Loop
            Set dict = Nothing
        
        Case PDF_ValueType.PDF_StreamData
            ' Note: PDF_ValueType.PDF_Stream is handled in PDF_Object
            ' this is only the stream ... endstream portion of a PDF stream
            tmpStr = GetLine(bytes, offset) ' skip stream, note dictionary previous read in should have a /Length value (passed in as meta)
            ' skip new line, data starts immediately after and may containe additional whitespace
            If bytes(offset) = 13 Then offset = offset + 1 ' vbCr
            If bytes(offset) = 10 Then offset = offset + 1 ' vbLf
            
            Dim dataBytes() As Byte
            If Meta Is Nothing Then Set Meta = New Dictionary
            If Meta.Exists("/Length") Then
                Dim byteLen As Long
                If IsObject(Meta.item("/Length").value) Then
                    ' Reference to actual /Length value, allowed but not common
                    Stop
                    ' TODO we need Dictionary of loaded objects or xrefTable to load value
                Else
                    ' /Length directly in file as number, usual
                    byteLen = CLng(Meta.item("/Length").value)
                End If
                ReDim dataBytes(0 To byteLen - 1)
                CopyBytes bytes, dataBytes, offset, 0, byteLen
                offset = offset + byteLen
                offset = SkipWhitespace(bytes, offset)
                tmpStr = GetWord(bytes, offset) ' endstream
            Else
                ' missing required key, as fallback read until an endstream line
                Debug.Print "Warning: PDF noncompliant, stream missing required /Length key"
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
            
            Set result = New pdfValue
            result.valueType = PDF_ValueType.PDF_StreamData
            result.value = dataBytes
            If Not IsMatch(Left(tmpStr, 9), "endstream") Then Stop ' error unexpected token found!
        
        Case PDF_ValueType.PDF_Reference, PDF_ValueType.PDF_Object
            ' Read three words at once
            Dim words(0 To 2) As String
            words(0) = GetWord(bytes, offset)
            offset = SkipWhitespace(bytes, offset)
            words(1) = GetWord(bytes, offset)
            offset = SkipWhitespace(bytes, offset)
            words(2) = GetWord(bytes, offset)
            offset = SkipWhitespace(bytes, offset)
            
            If IsMatch(words(2), "obj") Then
                Set result = NewValue(PDF_ValueType.PDF_Object, _
                    value:=GetValue(bytes, offset, contentStream:=contentStream, objRefCache:=objRefCache), _
                    objId:=CLng(words(0)), _
                    generation:=CLng(words(1)))
                
                ' load rest of object (a dictionary << >> and possibly stream) up to endobject
                If result.value.valueType = PDF_ValueType.PDF_Dictionary Then
                    Set Meta = result.value.value
                Else
                    Set Meta = Nothing
                End If
                
                ' get start of stream or endobject (skip if object is empty)
                Dim endObjOrStream As pdfValue
                If result.value.valueType = PDF_ValueType.PDF_EndOfObject Then
                    Debug.Print "Warning: empty object(" & result.id & "," & result.generation & ") - using null"
                    Set endObjOrStream = result.value
                    Set result.value = pdfValue.NullValue
                Else
                    Set endObjOrStream = GetValue(bytes, offset, Meta:=Meta, contentStream:=contentStream, objRefCache:=objRefCache)
                End If
                
                ' load stream, Note: stream object wraps & replaces initial dictionary .value
                If endObjOrStream.valueType = PDF_ValueType.PDF_StreamData Then
                    Dim stream As pdfValue: Set stream = New pdfValue
                    stream.valueType = PDF_ValueType.PDF_Stream
                    Dim streamObj As pdfStream: Set streamObj = New pdfStream
                    streamObj.Init result.value, endObjOrStream
                    Set stream.value = streamObj
                    Set streamObj = Nothing
                    Set result.value = stream
                    Set stream = Nothing
                    Set endObjOrStream = GetValue(bytes, offset, contentStream:=contentStream, objRefCache:=objRefCache)
                End If
                
                If endObjOrStream.valueType <> PDF_ValueType.PDF_EndOfObject Then
                    Stop ' Error: expected "endobj"
                End If
            Else ' Reference
                Set result = NewValue(PDF_ValueType.PDF_Reference, objId:=CLng(words(0)), generation:=CLng(words(1)))
            End If
            
            ' ensure same 'indirect reference object' instance returned for matching object id is always returned
            If (result.id <> 0) And (Not objRefCache Is Nothing) Then
                Dim objKey As String, refKey As String
                objKey = result.id & " " & result.generation ' object key = "id generation"
                refKey = objKey & " R" ' id generation R"
                Select Case result.valueType
                    Case PDF_ValueType.PDF_Object
                        ' ensure only 1 VBA instance of pdf obj
                        If objRefCache.Exists(objKey) Then
                            Set result = objRefCache(objKey) ' use previously created instance
                        Else
                            Set objRefCache(objKey) = result ' store instance for later use
                        End If
                        ' ensure only 1 VBA instance of pdf indirect obj reference
                        If objRefCache.Exists(refKey) Then
                            Set result.referenceObj = objRefCache(refKey)
                        Else
                            Set objRefCache(refKey) = result.referenceObj
                        End If
                    Case PDF_ValueType.PDF_Reference
                        ' ensure only 1 VBA instance of pdf indirect obj reference
                        If objRefCache.Exists(refKey) Then
                            Set result = objRefCache(refKey)
                        Else
                            Set objRefCache(refKey) = result
                        End If
                        ' Note: we don't preload obj as /Parent references can cause infinite loops & caller may not need obj yet
                        ' but if already loaded, ensure only 1 VBA instance of pdf obj
                        If objRefCache.Exists(objKey) Then
                            Set result.value = objRefCache(objKey) ' use previously created instance
                        End If
                    'Case Else
                        ' nothing to do
                End Select
            End If
            
        Case PDF_ValueType.PDF_Comment
            ' comment is everything until new line (strip % from beginning)
            Set result = pdfValue.NewValue(PDF_ValueType.PDF_Comment, Mid$(GetLine(bytes, offset), 2))
            
        Case PDF_ValueType.PDF_Trailer
            If Not IsMatch(GetWord(bytes, offset), "trailer") Then
                Stop ' Error: expected "trailer"
            End If
            Set result = New pdfValue
            result.valueType = valueType
            Set result.value = GetValue(bytes, offset, contentStream:=contentStream, objRefCache:=objRefCache)
            
        Case PDF_ValueType.PDF_EndOfArray, PDF_ValueType.PDF_EndOfDictionary, PDF_ValueType.PDF_EndOfObject, PDF_ValueType.PDF_EndOfStream
            tmpStr = GetWord(bytes, offset) ' Skip past end marker

        Case Else
            Stop ' Error: unexpected type
    End Select
    
    ' Skip trailing whitespace
    offset = SkipWhitespace(bytes, offset)
    
    ' set our value here as result may now be a different instance (singleton value)
    Set GetValue = result
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function

' Converts an ASC(char) value for a character representing a hex digit 0-F and returns numerical value
' If valid hexByte given, result is always in range of 0 to 15
Private Function HexDigitValue(ByVal hexByte As Byte) As Integer
    If hexByte >= 48 And hexByte <= 57 Then ' 0-9
        HexDigitValue = hexByte - 48
    ElseIf hexByte >= 65 And hexByte <= 70 Then ' A-F
        HexDigitValue = hexByte - 55
    ElseIf hexByte >= 97 And hexByte <= 102 Then ' a-f
        HexDigitValue = hexByte - 87
    End If
End Function

' Helper function to process string buffer with encoding detection
' converts length characters from buffer byte array to a String
' assumes buffer is either PdfDocEncoding, UTF-8, or UTF-16 BE/LE and converts to VBA Unicode String
' Note: buffer may be larger than length bytes, only first length bytes converted
Friend Function ProcessStringBuffer(ByRef buffer() As Byte, ByVal length As Long) As String
    ' return empty string if no data
    If length = 0 Then Exit Function
    
    ' with a byte buffer of values 0-255, we need to determine if
    ' encoding is pdfDocEncoding, UTF-8, or UTF-16 (previosly should be BE but now LE also used)
    ' UTF-8 has BOM bytes 239, 187 and 191
    ' Unicode UTF-16 has BOM bytes 255, 254 for BigEndian and 254, 255 for LittleEndian
    
    ' Note: to simplify and not access bytes outside range we get 1st 3 bytes or leave as 0
    ' to check for BOM marker
    Dim b1 As Byte, b2 As Byte, b3 As Byte
    If length >= 1 Then b1 = buffer(0)
    If length >= 2 Then b2 = buffer(1)
    If length >= 3 Then b3 = buffer(2)
    
    Dim buf() As Byte
    Dim ndx As Long
    If (b1 = 239) And (b2 = 187) And (b3 = 191) Then ' UTF-8 BOM
        ' strip BOM
        CopyBytes buffer, buf, 3, 0, length - 3
        ' convert to String from UTF8 encoded buffer
        ProcessStringBuffer = Utf8BytesToString(buf, length - 3)
    ElseIf (b1 = 254) And (b2 = 255) Then ' UTF-16BE BOM
        ' ignore trailing byte if odd number of bytes
        If length Mod 2 <> 0 Then length = length - 1
        Dim s As String
        For ndx = 2 To length - 1 Step 2
            Dim uChrVal As Long     ' may exceed MAXINT as we need up to 64K
            uChrVal = LShift(buffer(ndx), 8) + buffer(ndx + 1)
            s = s & ChrW(uChrVal)
        Next ndx
        ProcessStringBuffer = s
    ElseIf (b1 = 255) And (b2 = 254) Then ' UTF-16LE BOM
        ' strip BOM
        CopyBytes buffer, buf, 2, 0, length - 2
        ' VBA supports converting unicode byte array to string using 2 bytes (integer) per character
        ' *** TODO verify this if > ASCII
        ProcessStringBuffer = buf
    Else ' PDFDocEncoding
        ' Note: PDFDocEncoding has values 0-255 which correspond directly with 1st 256 Unicode values
        ' and all printable Unicode values in 0-255 directly correspond to Windows code page 1252 values
        ProcessStringBuffer = String(length, " ")  ' Pre-allocate string
        ' copy over bytes converted to 2 byte Unicode values (String characters)
        For ndx = 0 To length - 1
            Mid(ProcessStringBuffer, ndx + 1, 1) = ChrW(buffer(ndx))
        Next ndx
    End If
End Function

' Check if token is a known PDF operator
' Note: all known operators should be listed here as parsing content streams will likely fail otherwise
Private Function IsKnownOperator(ByVal tok As String, ByRef op As PdfTextOp) As Boolean
    IsKnownOperator = True
    op = PdfTextOp.opUnknown
    
    Select Case tok
        ' Text ops
        Case "TJ"
            op = PdfTextOp.opT_J
        Case "Tj"
            op = PdfTextOp.opTj
        Case "'"
            op = PdfTextOp.opQuote
        Case """"
            op = PdfTextOp.opDoubleQuote
        
        ' Text positioning
        Case "Td"
            op = PdfTextOp.opTd
        Case "TD"
            op = PdfTextOp.opT_D
        Case "Tm"
            op = PdfTextOp.opTm
        Case "T*"
            op = PdfTextOp.opTStar
        ' Text objects (begin/end text)
        Case "BT"
            op = PdfTextOp.opTextObjectBegin
        Case "ET"
            op = PdfTextOp.opTextObjectEnd
        ' Text state
        Case "Tf"
            op = PdfTextOp.opTf
        Case "Tc", "Tw", "Tz", "TL", "Tr", "Ts"
        
        ' Graphics state push / pop
        Case "q", "Q"
        ' Graphics operators
        Case "cm", "w", "J", "j", "M", "d", "ri", "i", "gs"
        ' Path construction
        Case "m", "l", "c", "v", "y", "h", "re"
        ' Path painting
        Case "S", "s", "f", "F", "f*", "B", "B*", "b", "b*", "n"
        ' Clipping paths
        Case "W", "W*"
        ' Marked content
        Case "MP", "DP", "BMC", "BDC", "EMC"
        ' backward compatibility / ignore errors on unknown operators (begin/end)
        Case "BX", "EX"
        ' XObject
        Case "Do"
        ' Type 3 Fonts
        Case "do", "d1"
        ' Color
        Case "CS", "cs", "SC", "SCN", "sc", "scn", "G", "g", "RG", "rg", "K", "k"
        ' Shading patterns
        Case "sh"
        ' Inline images
        Case "BI", "ID", "EI"
        Case Else
            IsKnownOperator = False
    End Select
End Function

