VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pdfStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' holds PDF data associated with a Stream value type, the Dictionary and Stream bytes
Option Explicit

Public stream_meta As pdfValue ' PDF_Dictionary
Public stream_data As pdfValue ' PDF_StreamData

' returns size of data(), but warns if differs from value of meta /Length property
Public Property Get Length() As Long
    ' Length is stored in meta
    ' it should be same as size of data
    If Meta.Exists("/Length") Then
        Length = CLng(Meta.Item("/Length"))
    End If
    
    On Error Resume Next
    Dim dataLen As Long
    dataLen = UBound(data) - LBound(data)
    If Err.Number <> 9 Then ' array out of bounds, i.e. not initialized
        If Length <> dataLen Then
            Debug.Print "Warning: data(0 to " & dataLen & ") array size does not match meta data /Length=" & Length & " value!"
            Length = dataLen ' use size of actual data, it coudl be modified
        End If
    End If
End Property

' this will add/update /Length property in meta and resize data()
Public Property Let Length(ByVal Count As Long)
    Meta("/Length") = Count ' add or replace
    
    On Error Resume Next
    Dim dataLen As Long
    dataLen = UBound(data) - LBound(data)
    On Error GoTo 0
    If Err.Number = 9 Then ' array out of bounds, i.e. not initialized
        ReDim data(0 To Count)
    ElseIf dataLen <> Count Then ' no need to adjust size if unchanged (ie just setting meta /Length property
        ReDim Preserve data(0 To Count) ' keep any data there, allows extending
    End If
End Property

' returns stream data as a Byte array
Public Property Get data() As Byte()
    data = stream_data.Value
End Property

' returns meta data as a VBA Dictionary object
Public Property Get Meta() As Dictionary
    Set Meta = stream_meta.Value
End Property

Public Sub Init(ByRef Meta As pdfValue, ByRef data As pdfValue)
    Set stream_meta = Meta
    Set stream_data = data
End Sub

Public Sub Class_Initialize()
    Set stream_meta = New pdfValue
    stream_meta.valueType = PDF_ValueType.PDF_Null
End Sub

Public Sub Class_Terminate()
    Set stream_meta = Nothing
    Set stream_data = Nothing
End Sub
