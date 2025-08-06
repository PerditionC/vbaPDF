Attribute VB_Name = "Test_pdfValue"
'====================================================================
'  Module: Test_pdfValue
'
'  Purpose:
'     unit-test suite for the pdfValue class.
'
'  How to run:
'     From the Immediate window  -  RunAll_pdfValue_Tests
'     A summary of passes/fails prints to the Immediate window.
'
'  Conventions:
'     - Every test sub is prefixed "Test_pdfValue_" to avoid conflicts.
'     - Helper assertions keep test code concise.
'====================================================================

Option Explicit

'--------------------------------------------------------------------
'              INTERNAL STATE FOR RUN COUNTERS
'--------------------------------------------------------------------
Private PassCount As Long
Private FailCount As Long

'--------------------------------------------------------------------
'              HELPER ASSERTION ROUTINES
'--------------------------------------------------------------------
' Asserts two Variant values are equal (uses = comparison)
Private Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, ByVal testName As String)
    If expected = actual Then
        Debug.Print "[PASS] " & testName
        PassCount = PassCount + 1
    Else
        Debug.Print "[FAIL] " & testName & _
                    "   Expected: " & expected & " ;  Got: " & actual
        FailCount = FailCount + 1
    End If
End Sub

' Asserts a Boolean condition is True
Private Sub AssertTrue(ByVal condition As Boolean, ByVal testName As String)
    If condition Then
        Debug.Print "[PASS] " & testName
        PassCount = PassCount + 1
    Else
        Debug.Print "[FAIL] " & testName & "   Condition evaluated False"
        FailCount = FailCount + 1
    End If
End Sub

' Asserts a Variant is Null (Variant subtype = Null)
Private Sub AssertIsNull(ByVal v As Variant, ByVal testName As String)
    AssertTrue IsNull(v), testName
End Sub

' Asserts an Object reference is Nothing
Private Sub AssertIsNothing(ByVal obj As Object, ByVal testName As String)
    AssertTrue obj Is Nothing, testName
End Sub


'====================================================================
'                           CORE TESTS
'====================================================================

'--------------------------------------------------
' 1) Constructor helpers - String
'--------------------------------------------------
' Verifies that CreateString correctly sets internal value & type
' Expected: .Value = supplied text, .ValueType = "String"
Public Sub Test_pdfValue_CreateString()
    Dim v As pdfValue: Set v = pdfValue.NewValue(PDF_String, "Hello PDF")
    AssertEqual "Hello PDF", v.value, "CreateString stores value"
    AssertEqual PDF_String, v.valueType, "CreateString sets ValueType"
End Sub

'--------------------------------------------------
' 2) Constructor helpers - Number
'--------------------------------------------------
Public Sub Test_pdfValue_CreateNumber()
    Dim v As pdfValue: Set v = pdfValue.NewValue(PDF_Real, 123.45)
    AssertEqual 123.45, v.value, "CreateNumber stores numeric value"
    AssertEqual PDF_Real, v.valueType, "CreateNumber sets ValueType"
    Set v = pdfValue.NewValue(PDF_Integer, 123.45)
    AssertEqual 123, v.value, "CreateNumber stores numeric value"
    AssertEqual PDF_Integer, v.valueType, "CreateNumber sets ValueType"
End Sub

'--------------------------------------------------
' 3) Constructor helpers - Boolean + ToString tokens
'--------------------------------------------------
Public Sub Test_pdfValue_CreateBoolean_And_ToString()
    AssertEqual "true", pdfValue.NewValue(PDF_Boolean, True).ToString, "Boolean TRUE token"
    AssertEqual "false", pdfValue.NewValue(PDF_Boolean, False).ToString, "Boolean FALSE token"
End Sub

'--------------------------------------------------
' 4) Constructor helpers - pdf Null value
'--------------------------------------------------
Public Sub Test_pdfValue_CreateNull()
    Dim v As pdfValue: Set v = pdfValue.NewValue(PDF_Null)
    AssertIsNull v.value, "Null value is Variant Null"
    AssertEqual PDF_Null, v.valueType, "CreateNull sets ValueType"
    AssertEqual "null", v.ToString, "Null ToString token"
    AssertTrue v Is pdfValue.NullValue, "null is interned"
End Sub

'--------------------------------------------------
' 5) Constructor helpers - Name (auto-slash & pre-slashed)
'--------------------------------------------------
Public Sub Test_pdfValue_CreateName_SlashHandling()
    ' Name without leading slash â–¸ slash is added
    Dim v1 As pdfValue: Set v1 = pdfValue.NewValue(PDF_Name, "Author")
    AssertEqual "/Author", v1.value, "Auto-slash added"

    ' Name already containing slash â–¸ unchanged
    Dim v2 As pdfValue: Set v2 = pdfValue.NewValue(PDF_Name, "/Title")
    AssertEqual "/Title", v2.value, "Pre-slashed name unchanged"
End Sub

'--------------------------------------------------
' 6) Constructor helpers - Reference object
'--------------------------------------------------
Public Sub Test_pdfValue_CreateReference()
    Dim r As pdfValue: Set r = pdfValue.NewValue(PDF_Reference, , 10, 2)
    AssertEqual 10, r.id, "Reference stores object ID"
    AssertEqual 2, r.generation, "Reference stores generation"
    AssertEqual "10 2 R", r.ToString, "Reference ToString format"

    ' validate shared reference object instances
    Dim obj As pdfValue
    Set obj = pdfValue.NewValue(PDF_Object, r, objId:=5, generation:=1)

    Dim ref1 As pdfValue, ref2 As pdfValue
    Set ref1 = obj.referenceObj
    Set ref2 = obj.referenceObj
    AssertTrue (ref1 Is ref2), "References refer to same instance"
End Sub

'--------------------------------------------------
' 7) Array creation, mutation, and serialisation
'--------------------------------------------------
Public Sub Test_pdfValue_Array_CreateAddRender()
    ' Empty array â–¸ Count = 0 â–¸ ToString "[]"
    Dim a As pdfValue: Set a = pdfValue.NewValue(PDF_Array)
    AssertEqual 0, a.value.count, "New array starts empty"
    AssertEqual "[]", Replace(a.ToString, " ", ""), "Empty array render"

    ' Add three mixed items and verify
    a.Add pdfValue.NewValue(PDF_Integer, 1)
    a.Add pdfValue.NewValue(PDF_Real, 2)
    a.Add pdfValue.NewValue(PDF_String, "X")
    AssertEqual 3, a.value.count, "Add increments count"
    AssertEqual "[ 1 2.0 (X) ]", a.ToString, "Array ToString with mixed items"
End Sub

'--------------------------------------------------
' 8) Dictionary creation, mutation, nested render
'--------------------------------------------------
Public Sub Test_pdfValue_Dictionary_CreateSetGetRender()
    Dim d As pdfValue: Set d = pdfValue.NewValue(PDF_Dictionary)

    ' Empty dict¸ ToString "<< >>"
    AssertTrue InStr(d.ToString, "<<") = 1 And InStr(d.ToString, ">>") > 0, _
               "Empty dictionary render"

    ' Set two keys
    d.SetDictionaryValue "/Type", pdfValue.GetInternedName("Catalog")
    d.SetDictionaryValue "/Count", pdfValue.NewValue(PDF_Integer, 5)

    ' Retrieve existing key
    Dim got As pdfValue: Set got = d.GetDictionaryValue("/Type")
    AssertEqual "/Catalog", got.value, "GetDictionaryValue existing key"

    ' Retrieve missing key – Expect Nothing
    Dim missing As Variant: Set missing = d.GetDictionaryValue("/Missing")
    AssertTrue (IsObject(missing) And missing Is Nothing), _
               "GetDictionaryValue missing key returns Nothing"

    ' Render includes both entries
    Dim txt As String: txt = d.ToString
    AssertTrue InStr(txt, "/Type /Catalog") > 0, "Dictionary render contains Type"
    AssertTrue InStr(txt, "/Count 5") > 0, "Dictionary render contains Count"
End Sub

'--------------------------------------------------
' Exists for Dictionaries and dictionary objects
'--------------------------------------------------
Public Sub Test_pdfValue_Dictionary_Exists()
    Dim dict As pdfValue
    Set dict = pdfValue.NewValue(PDF_Dictionary)
    dict.SetDictionaryValue "/Type", pdfValue.GetInternedName("Page")

    AssertTrue dict.Exists("Type"), "PDF_Dictionary.Exists ""Type"" without leading slash"
    AssertTrue dict.Exists("/Type"), "PDF_Dictionary.Exists ""/Type"" with leading slash"
    AssertTrue Not dict.Exists("Missing"), "PDF_Dictionary.Exists is False for key that does not exist"

    ' Works with PDF objects containing dictionaries
    Dim obj As New pdfValue
    obj.valueType = PDF_Object
    Set obj.value = dict
    AssertTrue obj.Exists("Type"), "PDF_Dictionary.Exists ""Type"" for embedded PDF_Dictionary in PDF_Object"
End Sub

'--------------------------------------------------
' Get/Set Dictionary values
'--------------------------------------------------
Public Sub Test_pdfValue_Dictionary_GetSetValue()
    Dim dict As pdfValue
    Set dict = pdfValue.NewValue(PDF_Dictionary)
    dict.SetDictionaryValue "/Type", pdfValue.GetInternedName("Page")

    Dim typeVal As pdfValue
    Set typeVal = dict.GetDictionaryValue("Type")
    AssertEqual typeVal.ToString(), "/Page", "GetDictionaryValue(SetDictionaryValue) match"
End Sub

'--------------------------------------------------
' 9) Nested structures (dict inside array)
'--------------------------------------------------
Public Sub Test_pdfValue_NestedStructureRender()
    Dim arr As pdfValue: Set arr = pdfValue.NewValue(PDF_Array)
    Dim inner As pdfValue: Set inner = pdfValue.NewValue(PDF_Dictionary)
    inner.SetDictionaryValue "/K", pdfValue.NewValue(PDF_String, "V")
    arr.Add inner
    AssertTrue InStr(arr.ToString, "/K (V)") > 0, "Nested dictionary rendered"
End Sub

'--------------------------------------------------
'10) ID / Generation mutability defaults
'--------------------------------------------------
Public Sub Test_pdfValue_DefaultInitialState_And_FieldSet()
    Dim v As New pdfValue
    AssertEqual 0, v.id, "Default ID = 0"
    AssertEqual 0, v.generation, "Default Generation = 0"
    AssertEqual PDF_Null, v.valueType, "Default ValueType = pdfNull"

    ' Mutate fields
    v.id = 77: v.generation = 3
    AssertEqual 77, v.id, "ID mutable"
    AssertEqual 3, v.generation, "Generation mutable"
End Sub

'--------------------------------------------------
'11) Helper misuse (array/dict ops on wrong type) must not crash
'--------------------------------------------------
Public Sub Test_pdfValue_HelperMisuseSafety()
    On Error GoTo errHandler
    Dim num As pdfValue: Set num = pdfValue.NewValue(PDF_Integer, 9)
    Dim v As pdfValue
    Set v = pdfValue.NewValue(PDF_Integer, 1)
    On Error GoTo ErrExpected
    num.Add v ' Add should be throw an error
    On Error GoTo errHandler
    num.SetDictionaryValue "/A", pdfValue.NewValue(PDF_Integer, 1)  ' Should be ignored
    Dim res: Set res = num.GetDictionaryValue("/A")   ' Should return Nothing
    AssertTrue (IsObject(res) And res Is Nothing), _
               "Misuse returns Nothing"
    AssertTrue True, "No runtime error on helper misuse"
    Exit Sub
ErrExpected:
    If Err.Number = 13 Then Resume Next
errHandler:
    AssertTrue False, "Helper misuse raised unexpected error: " & Err.Description
End Sub

'--------------------------------------------------
'12) Unknown / unsupported ValueType - ToString returns empty
'    (Regression guard if enum extended in future)
'--------------------------------------------------
Public Sub Test_pdfValue_ToString_UnknownType()
    Dim v As New pdfValue
    v.valueType = 999        ' set to an invalid / future enum value
    AssertEqual "", v.ToString, "Unknown ValueType - empty string"
End Sub

'--------------------------------------------------
' simple ToString examples
'--------------------------------------------------
Public Sub Test_pdfValue_ToString()
    Dim obj As pdfValue
    AssertEqual pdfValue.NewValue(PDF_Null).ToString, "null", "ToString - null"
    AssertEqual pdfValue.NewValue(PDF_Boolean, True).ToString, "true", "ToString - true"
    AssertEqual pdfValue.NewValue(PDF_Boolean, False).ToString, "false", "ToString - false"
    AssertEqual pdfValue.NewValue(PDF_Integer, 42&).ToString, "42", "ToString - integer"
    AssertEqual pdfValue.NewValue(PDF_Real, 3.14159).ToString, "3.14159", "ToString - double"
    Set obj = pdfValue.NewValue(PDF_String, "Hello World")
    AssertEqual obj.ToString, "(Hello World)", "ToString - string (text)"
    obj.flags = obj.flags Or flgBinary
    AssertEqual obj.ToString, "<48 65 6C 6C 6F 20 57 6F 72 6C 64>", "ToString - string <hex>"
    AssertEqual pdfValue.NewValue(PDF_Name, "Type").ToString, "/Type", "ToString - /Name"
    AssertEqual pdfValue.NewArray(1, 2, 3).ToString, "[ 1 2 3 ]", "ToString - array of integers"
    AssertEqual pdfValue.NewValue(PDF_Array, pdfValue.NewArray(1, 2, 3)).ToString, "[ [ 1 2 3 ] ]", "ToString - array of array of integers"
    Set obj = pdfValue.NewValue(PDF_Dictionary)
    obj.SetDictionaryValue "Type", pdfValue.GetInternedName("Page")
    AssertEqual obj.ToString, "<<" & vbLf & "/Type /Page" & vbLf & ">>" & vbLf, "ToString - dictionary of /Name"
    AssertEqual pdfValue.NewValue(PDF_Reference, objId:=5).ToString, "5 0 R", "ToString - indirect object reference"
End Sub

'--------------------------------------------------
'13) Enumeration values (sanity check: first few constants)
'--------------------------------------------------
Public Sub Test_pdfValue_EnumValues()
    ' first 9 types correspond to core pdf value types
    AssertEqual 0, PDF_Null, "Enum PDF_Null = 0"
    AssertEqual 1, PDF_Boolean, "Enum PDF_Boolean = 1"
    AssertEqual 2, PDF_Integer, "Enum PDF_Integer = 2"
    AssertEqual 7, PDF_Dictionary, "Enum PDF_Dictionary = 2"
End Sub

'--------------------------------------------------
' /Name interning
'--------------------------------------------------
Public Sub Test_pdfValue_InternName()
    Dim typeVal1 As pdfValue, typeVal2 As pdfValue
    Set typeVal1 = pdfValue.GetInternedName("Type")
    Set typeVal2 = pdfValue.GetInternedName("/Type")
    AssertTrue (typeVal1 Is typeVal2), "GetInternedName values ""Type"" and ""/Type"" return same instance"
End Sub

'--------------------------------------------------
'14) simple pdf array creation
'--------------------------------------------------
Public Sub Test_pdfValue_NewArray()
    Dim v As New pdfValue
    Dim o As New Dictionary
    Set v = pdfValue.NewArray(1, 2.1, "A String", "/MyName", o, Null)
    AssertEqual 6, v.count, "6 items added"
    AssertEqual 1, v.GetArrayValue(1).value, "1st array item is 1"
    AssertEqual 2.1, v.GetArrayValue(2).value, "2nd array item is Double 2.1"
    AssertEqual PDF_ValueType.PDF_String, v.GetArrayValue(3).valueType, "is a pdf string"
    AssertEqual PDF_ValueType.PDF_Name, v.GetArrayValue(4).valueType, "is a pdf name"
    AssertEqual PDF_ValueType.PDF_Dictionary, v.GetArrayValue(5).valueType, "object is pdf dictionary"
    AssertTrue pdfValue.NullValue Is v.GetArrayValue(6), "Null becomes pdf null"
End Sub


'--------------------------------------------------
'15) alternative simple pdf array creation
'--------------------------------------------------
Public Sub Test_pdfValue_NewArrayAlt()
    Dim v As New pdfValue
    Dim o As New Dictionary
    Dim myarray As Variant
    myarray = Array(1, 2.1, "A String", "/MyName", o, Null)
    Set v = pdfValue.NewArray(myarray)
    AssertEqual 6, v.count, "6 items added"
    AssertEqual 1, v.GetArrayValue(1).value, "1st array item is 1"
    AssertEqual 2.1, v.GetArrayValue(2).value, "2nd array item is Double 2.1"
    AssertEqual PDF_ValueType.PDF_String, v.GetArrayValue(3).valueType, "is a pdf string"
    AssertEqual PDF_ValueType.PDF_Name, v.GetArrayValue(4).valueType, "is a pdf name"
    AssertEqual PDF_ValueType.PDF_Dictionary, v.GetArrayValue(5).valueType, "object is pdf dictionary"
    AssertTrue pdfValue.NullValue Is v.GetArrayValue(6), "Null becomes pdf null"
End Sub


'--------------------------------------------------
'16) simple pdf array addition
'--------------------------------------------------
Public Sub Test_pdfValue_AddArray()
    Dim v As New pdfValue
    Dim o As New Dictionary
    Set v = pdfValue.NewValue(PDF_Array)
    v.Add 1, 2.1, "A String", "/MyName", o, Null
    AssertEqual 6, v.count, "6 items added"
    AssertEqual 1, v.GetArrayValue(1).value, "1st array item is 1"
    AssertEqual 2.1, v.GetArrayValue(2).value, "2nd array item is Double 2.1"
    AssertEqual PDF_ValueType.PDF_String, v.GetArrayValue(3).valueType, "is a pdf string"
    AssertEqual PDF_ValueType.PDF_Name, v.GetArrayValue(4).valueType, "is a pdf name"
    AssertEqual PDF_ValueType.PDF_Dictionary, v.GetArrayValue(5).valueType, "object is pdf dictionary"
    AssertTrue pdfValue.NullValue Is v.GetArrayValue(6), "Null becomes pdf null"
End Sub

'--------------------------------------------------
' test deep cloning
'--------------------------------------------------
Public Sub Test_pdfValue_Clone()
    ' Clone a simple value
    Dim original As pdfValue
    Set original = pdfValue.NewValue(PDF_String, "Hello World")
    Dim copy As pdfValue
    Set copy = original.Clone()
    AssertEqual original.value, copy.value, "Cloned strings match"
    AssertTrue Not copy Is original, "Cloned objects are different instances"

    ' Clone a complex structure
    Dim originalDict As pdfValue
    Set originalDict = pdfValue.NewTypedDictionary("Page")
    originalDict.SetDictionaryValue "/MediaBox", pdfValue.NewArray(0, 0, 612, 792)
    Dim clonedDict As pdfValue
    Set clonedDict = originalDict.Clone()
    AssertEqual originalDict.count, clonedDict.count, "Cloned dictionaries same count of items"
    AssertTrue Not clonedDict Is originalDict, "Clone dictionary is different instance"
    Dim innerArray1 As pdfValue, innerArray2 As pdfValue
    Set innerArray1 = originalDict.GetDictionaryValue("/MediaBox")
    Set innerArray2 = clonedDict.GetDictionaryValue("MediaBox")
    AssertEqual innerArray1.count, innerArray2.count, "Cloned nested arrays same count of items"
    AssertEqual innerArray1.GetArrayValue(3).value, innerArray2.GetArrayValue(3).value, "Cloned nested arrays have same values"
    AssertEqual 612, innerArray2.GetArrayValue(3).value, "Cloned nested array has expected values"
    AssertTrue Not innerArray2 Is innerArray1, "Cloned nested array is different instance"
    

    ' Changes to clone don't affect original
    clonedDict.SetDictionaryValue "/Rotate", pdfValue.NewValue(PDF_Integer, 90)
    ' originalDict remains unchanged
    AssertTrue clonedDict.Exists("/Rotate") And Not originalDict.Exists("/Rotate"), _
        "Key in cloned but not in original - changes don't effect original"

    ' Clone with object reference cache for consistency
    Dim objCache As New Dictionary
    Dim obj1 As pdfValue: Set obj1 = pdfValue.NewValue(PDF_Object, clonedDict, 1, 0)
    Dim obj2 As pdfValue: Set obj2 = pdfValue.NewValue(PDF_Object, innerArray2, 2, 0)
    Dim obj3 As pdfValue: Set obj3 = pdfValue.NewValue(PDF_Object, pdfValue.NewValue(PDF_String, "Shared object"), 3, 99)
    obj1.GetDictionaryValue("/MediaBox").Add obj3 ' obj3 is now item(5) of innerArray2
    obj2.Add obj3 ' obj3 is now also item(6) of innerArray2
    AssertEqual innerArray2.count, 6, "Shared string object added twice to same array"
    ' Shared references between obj1 and obj2 are preserved
    AssertTrue obj3 Is innerArray2.GetArrayValue(6), "objects correctly shared when added"
    AssertTrue innerArray2.GetArrayValue(5) Is innerArray2.GetArrayValue(6), "objects correctly shared"
    Dim cloneObj1 As pdfValue: Set cloneObj1 = obj1.Clone(objCache)
    Dim cloneObj2 As pdfValue: Set cloneObj2 = obj2.Clone(objCache)
    Set obj1 = cloneObj1.GetDictionaryValue("/MediaBox").GetArrayValue(6)
    Set obj2 = cloneObj2.GetArrayValue(6)
    obj1.generation = 33
    AssertTrue obj2 Is obj1, "objects correctly shared/interned with respect to same obj reference cache"
    Set obj1 = cloneObj1.GetDictionaryValue("/MediaBox").GetArrayValue(5)
    AssertTrue obj2 Is obj1, "objects correctly shared/interned cloned nested item with respect to same obj reference cache"
End Sub

'----------------------------------------------------
'17)+ check our byte array to string text conversions
'----------------------------------------------------

' HELPER FOR CHECKING RESULTS
Private Sub CheckStringResult(ByVal result As String, ByVal expectedChars As Variant, ByVal testPrefix As String)
    ' Check string length
    AssertEqual UBound(expectedChars) - LBound(expectedChars) + 1, Len(result), testPrefix & " - String length"
    
    ' Check individual Unicode values
    Dim i As Long
    For i = LBound(expectedChars) To UBound(expectedChars)
        If i - LBound(expectedChars) + 1 <= Len(result) Then
            AssertEqual expectedChars(i), AscW(Mid(result, i - LBound(expectedChars) + 1, 1)), _
                       testPrefix & " - Char " & (i - LBound(expectedChars) + 1) & " Unicode value"
        End If
    Next i
End Sub

' TEST FUNCTIONS
Private Sub TestPdfDocEncoding()
    Debug.Print "--- Test 1: PdfDocEncoding ---"
    
    Dim testBuffer() As Byte
    Dim result As String
    Dim expectedChars As Variant
    
    ' Test string: "Aé" (A=65, é=233) - covers 0-127 and 128-255 ranges
    expectedChars = Array(65, 233)  ' A, é
    
    ReDim testBuffer(0 To 1)
    testBuffer(0) = 65   ' 'A' (0-127 range)
    testBuffer(1) = 233  ' 'é' (128-255 range)
    
    result = pdfValue.ProcessStringBuffer(testBuffer, 2)
    CheckStringResult result, expectedChars, "PdfDocEncoding"
    
    Debug.Print ""
End Sub

Private Sub TestUTF8WithBOM()
    Debug.Print "--- Test 2: UTF-8 encoding with BOM ---"
    
    Dim testBuffer() As Byte
    Dim result As String
    Dim expectedChars As Variant
    
    ' Test 1: Same final string as PdfDocEncoding - "Aé"
    expectedChars = Array(65, 233)  ' A, é
    
    ReDim testBuffer(0 To 5)
    ' UTF-8 BOM: EF BB BF
    testBuffer(0) = &HEF
    testBuffer(1) = &HBB
    testBuffer(2) = &HBF
    ' "Aé" in UTF-8: A(41) é(C3 A9)
    testBuffer(3) = 65   ' 'A'
    testBuffer(4) = &HC3 ' First byte of UTF-8 é
    testBuffer(5) = &HA9 ' Second byte of UTF-8 é
    
    result = pdfValue.ProcessStringBuffer(testBuffer, 6)
    CheckStringResult result, expectedChars, "UTF-8 Test 1"
    
    ' Test 2: Extended Unicode - "A€?" (A=65, €=8364, ?=8721)
    ' Covers 0-127, 128-255, and >255 ranges
    expectedChars = Array(65, 8364, 8721)  ' A, €, ?
    
    ReDim testBuffer(0 To 11)
    ' UTF-8 BOM: EF BB BF
    testBuffer(0) = &HEF
    testBuffer(1) = &HBB
    testBuffer(2) = &HBF
    ' "A€?" in UTF-8: A(41) €(E2 82 AC) ?(E2 88 91)
    testBuffer(3) = 65   ' 'A'
    testBuffer(4) = &HE2 ' First byte of UTF-8 €
    testBuffer(5) = &H82 ' Second byte of UTF-8 €
    testBuffer(6) = &HAC ' Third byte of UTF-8 €
    testBuffer(7) = &HE2 ' First byte of UTF-8 ?
    testBuffer(8) = &H88 ' Second byte of UTF-8 ?
    testBuffer(9) = &H91 ' Third byte of UTF-8 ?
    
    result = pdfValue.ProcessStringBuffer(testBuffer, 10)
    CheckStringResult result, expectedChars, "UTF-8 Test 2"
    
    Debug.Print ""
End Sub

Private Sub TestUTF16LittleEndian()
    Debug.Print "--- Test 3: UTF-16 Little Endian ---"
    
    Dim testBuffer() As Byte
    Dim result As String
    Dim expectedChars As Variant
    
    ' Test 1: Same final string as PdfDocEncoding - "Aé"
    expectedChars = Array(65, 233)  ' A, é
    
    ReDim testBuffer(0 To 5)
    ' UTF-16 LE BOM: FF FE
    testBuffer(0) = &HFF
    testBuffer(1) = &HFE
    ' "Aé" in UTF-16 LE: A(41 00) é(E9 00)
    testBuffer(2) = 65   ' 'A' low byte
    testBuffer(3) = &H0  ' 'A' high byte
    testBuffer(4) = 233  ' 'é' low byte
    testBuffer(5) = &H0  ' 'é' high byte
    
    result = pdfValue.ProcessStringBuffer(testBuffer, 6)
    CheckStringResult result, expectedChars, "UTF-16 LE Test 1"
    
    ' Test 2: Extended Unicode - "A€?" (A=65, €=8364, ?=8721)
    expectedChars = Array(65, 8364, 8721)  ' A, €, ?
    
    ReDim testBuffer(0 To 7)
    ' UTF-16 LE BOM: FF FE
    testBuffer(0) = &HFF
    testBuffer(1) = &HFE
    ' "A€?" in UTF-16 LE: A(41 00) €(AC 20) ?(11 22)
    testBuffer(2) = 65   ' 'A' low byte
    testBuffer(3) = &H0  ' 'A' high byte
    testBuffer(4) = &HAC ' '€' low byte (8364 = 20AC hex)
    testBuffer(5) = &H20 ' '€' high byte
    testBuffer(6) = &H11 ' '?' low byte (8721 = 2211 hex)
    testBuffer(7) = &H22 ' '?' high byte
    
    result = pdfValue.ProcessStringBuffer(testBuffer, 8)
    CheckStringResult result, expectedChars, "UTF-16 LE Test 2"
    
    Debug.Print ""
End Sub

Private Sub TestUTF16BigEndian()
    Debug.Print "--- Test 4: UTF-16 Big Endian ---"
    
    Dim testBuffer() As Byte
    Dim result As String
    Dim expectedChars As Variant
    
    ' Test 1: Same final string as PdfDocEncoding - "Aé"
    expectedChars = Array(65, 233)  ' A, é
    
    ReDim testBuffer(0 To 5)
    ' UTF-16 BE BOM: FE FF
    testBuffer(0) = &HFE
    testBuffer(1) = &HFF
    ' "Aé" in UTF-16 BE: A(00 41) é(00 E9)
    testBuffer(2) = &H0  ' 'A' high byte
    testBuffer(3) = 65   ' 'A' low byte
    testBuffer(4) = &H0  ' 'é' high byte
    testBuffer(5) = 233  ' 'é' low byte
    
    result = pdfValue.ProcessStringBuffer(testBuffer, 6)
    CheckStringResult result, expectedChars, "UTF-16 BE Test 1"
    
    ' Test 2: Extended Unicode - "A€?" (A=65, €=8364, ?=8721)
    expectedChars = Array(65, 8364, 8721)  ' A, €, ?
    
    ReDim testBuffer(0 To 7)
    ' UTF-16 BE BOM: FE FF
    testBuffer(0) = &HFE
    testBuffer(1) = &HFF
    ' "A€?" in UTF-16 BE: A(00 41) €(20 AC) ?(22 11)
    testBuffer(2) = &H0  ' 'A' high byte
    testBuffer(3) = 65   ' 'A' low byte
    testBuffer(4) = &H20 ' '€' high byte (8364 = 20AC hex)
    testBuffer(5) = &HAC ' '€' low byte
    testBuffer(6) = &H22 ' '?' high byte (8721 = 2211 hex)
    testBuffer(7) = &H11 ' '?' low byte
    
    result = pdfValue.ProcessStringBuffer(testBuffer, 8)
    CheckStringResult result, expectedChars, "UTF-16 BE Test 2"
    
    Debug.Print ""
End Sub



'====================================================================
'                        TEST RUNNER
'====================================================================
Public Sub RunAll_pdfValue_Tests()
    PassCount = 0: FailCount = 0
    Debug.Print "========== pdfValue Test Suite =========="

    ' List every test here (easy to comment out temporarily)
    Test_pdfValue_CreateString
    Test_pdfValue_CreateNumber
    Test_pdfValue_CreateBoolean_And_ToString
    Test_pdfValue_CreateNull
    Test_pdfValue_CreateName_SlashHandling
    Test_pdfValue_CreateReference
    Test_pdfValue_Array_CreateAddRender
    Test_pdfValue_Dictionary_CreateSetGetRender
    Test_pdfValue_Dictionary_Exists
    Test_pdfValue_Dictionary_GetSetValue
    Test_pdfValue_NestedStructureRender
    Test_pdfValue_DefaultInitialState_And_FieldSet
    Test_pdfValue_HelperMisuseSafety
    Test_pdfValue_ToString_UnknownType
    Test_pdfValue_ToString
    Test_pdfValue_EnumValues
    Test_pdfValue_InternName
    Test_pdfValue_NewArray
    Test_pdfValue_NewArrayAlt
    Test_pdfValue_AddArray
    Test_pdfValue_Clone

    ' PdfDocEncoding (direct byte-to-Unicode mapping)
    TestPdfDocEncoding
    ' UTF-8 encoding with BOM
    TestUTF8WithBOM
    ' UTF-16 Little Endian
    TestUTF16LittleEndian
    ' UTF-16 Big Endian
    TestUTF16BigEndian

    Debug.Print "========== Summary  ======================"
    Debug.Print "  Passed: " & PassCount
    Debug.Print "  Failed: " & FailCount
    Debug.Print "=========================================="
End Sub

