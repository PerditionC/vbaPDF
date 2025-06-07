Attribute VB_Name = "pdfValues"
' creates and returns pdfValue objects for given input
Option Explicit

' returns a /Name as a pdfValue obj
Function pdfNameObj(ByVal name As String) As pdfValue
#If False Then
    Set pdfNameObj = New pdfValue
    pdfNameObj.valueType = PDF_ValueType.PDF_Name
    pdfNameObj.Value = name
#Else
    Set pdfNameObj = pdfValueObj(name, "/Name")
#End If
End Function


' returns value as a pdfValue obj
' Note: if value is String then valueType can be used if want a PDF_Name or PDF_Trailer object instead of PDF_String
' Currently a Dictionary returns a PDF_Dictionary and a Collection returns as PDF_Array, no other objects are supported!
Function pdfValueObj(ByRef Value As Variant, Optional ByRef valueType As String = vbNullString) As pdfValue
    Dim obj As pdfValue
    Set obj = New pdfValue
    Select Case VarType(Value)
        Case vbLong, vbInteger
            obj.valueType = PDF_ValueType.PDF_Integer
            obj.Value = CLng(Value)
        Case vbSingle, vbDouble
            obj.valueType = PDF_ValueType.PDF_Real
            obj.Value = CDbl(Value)
        Case vbBoolean
            obj.valueType = PDF_ValueType.PDF_Boolean
            obj.Value = CBool(Value)
        Case vbString
            Select Case valueType
                Case "/Name"
                    obj.valueType = PDF_ValueType.PDF_Name
                    obj.Value = Value
                Case Else
                    obj.valueType = PDF_ValueType.PDF_String
                    obj.Value = Value.id
                    obj.generation = Value.generation
            End Select
        Case vbObject
            Select Case TypeName(Value)
                Case "Dictionary"
                    obj.valueType = PDF_ValueType.PDF_Dictionary
                    Set obj.Value = Value
                Case "Collection"
                    obj.valueType = PDF_ValueType.PDF_Array
                    Set obj.Value = Value
                Case "pdfValue"
                    Select Case valueType
                        Case "/Trailer"
                            obj.valueType = PDF_ValueType.PDF_Trailer
                            Set obj.Value = Value
                        Case "/Reference"
                            obj.valueType = PDF_ValueType.PDF_Reference
                            obj.Value = Value.id
                            obj.generation = Value.generation
                        Case Else
                            Stop ' ???
                    End Select
                Case Else
                    Stop ' ???
            End Select
    End Select
    Set pdfValueObj = obj
End Function


' returns a Collection as a pdfValue array [] obj
Function pdfArrayObj(ByRef Items As Collection) As pdfValue
#If False Then
    Set pdfArrayObj = New pdfValue
    pdfArrayObj.valueType = PDF_ValueType.PDF_Array
    Set pdfArrayObj.Value = Items
#Else
    Set pdfArrayObj = pdfValueObj(Items)
#End If
End Function


