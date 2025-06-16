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

' stream data may be compressed, on first access we to uncompressed data we decompress and cache results
Private m_udata() As Byte
Private decompressed As Boolean


' returns size of data(), but warns if differs from value of meta /Length property
Public Property Get Length() As Long
    ' Length is stored in meta
    ' it should be same as size of data
    If Meta.Exists("/Length") Then
        Length = CLng(Meta.Item("/Length").value)
    End If
    
    On Error Resume Next
    Dim dataLen As Long
    dataLen = UBound(data) - LBound(data) + 1
    If Err.Number <> 9 Then ' array out of bounds, i.e. not initialized
        If Length <> dataLen Then
            Debug.Print "Warning: data(0 to " & dataLen - 1 & ") array size does not match meta data /Length=" & Length & " value!"
            Length = dataLen ' use size of actual data, it coudl be modified
        End If
    End If
End Property

' this will add/update /Length property in meta and resize data()
Public Property Let Length(ByVal Count As Long)
    Meta("/Length").value = Count ' add or replace
    
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

' returns stream data [as stored, ie. compressed] as a Byte array
Public Property Get data() As Byte()
    data = stream_data.value
End Property

' returns stream data in uncompressed form as Byte array
Public Property Get udata() As Byte()
    If decompressed Then
        udata = m_udata
    Else
        ' get the uncompressed (un-/Filter'd) data if compressed (/Filter'd)
        If Meta.Exists("/Filter") Then
            ' TODO support all /Filter types
            Dim filter As String
            filter = Meta.Item("/Filter").value
            Select Case LCase(filter)
                Case "/flatedecode"
                    Dim startIndex As Long: startIndex = 2 ' skip past zlib wrapper to raw Deflate data
                    Dim outSize As Long ' as count of bytes -- should be 0 on entry so will be size on output
                    Dim estBufSize As Long ' if known, uncompressed size
                    If Meta.Exists("/DL") Then estBufSize = CLng(Meta.Item("/DL").value) ' only a hint
                    Erase m_udata
                    
                    ' Note: libdeflate_inflate arguments are byVal, we expect them be byRef
                    'libdeflate_inflate objStream.data, startIndex, m_udata, outSize
                    If Not inflate2(data, m_udata, startIndex, outSize, estBufSize) Then
                        Debug.Print "Error decompressing!"
                        Stop
                    End If
                    
                    ' do we need to decode uncompressed stream?
                    Dim predictor As Long, columns As Long
                    ' set default values if not otherwise specified
                    predictor = 1: columns = 1
                    If Meta.Exists("/DecodeParms") Then
                        Dim pdfV As pdfValue
                        Set pdfV = Meta.Item("/DecodeParms")
                        Dim decodeParms As Dictionary
                        Set decodeParms = pdfV.value
                        If decodeParms.Exists("/Predictor") Then
                            Set pdfV = decodeParms.Item("/Predictor")
                            predictor = pdfV.value
                        End If
                        ' should only be supplied if predictor > 1, but we can load value regardless, only used if predictor > 1
                        If decodeParms.Exists("/Columns") Then
                            Set pdfV = decodeParms.Item("/Columns")
                            columns = pdfV.value
                        End If
                        Set pdfV = Nothing
                    End If
                    ' if predictor > 1 then we need to reverse differencing done prior to compression/encoding
                    Dim rowOffset As Long, rowIndex As Long
                    Select Case predictor
                        Case 1 ' default no prediction
                            ' nothing to do
                        Case 2 ' TIFF predictor
                            Stop ' not implemented
                        
                        ' Note: regardless of predictor specified, if PNG filter then each row should have a predicator tag, need not match predictor value
                        ' 10 through 15 are defined in PNG RFC 2083 specification, see Chapter 6, Filters
                        Case 10, 11, 12, 13, 14, 15
                            ' actual data is smaller, we need to buffer our data
                            Dim buffer() As Byte
                            ReDim buffer(0 To UBound(m_udata))
                            Dim bufferIndex As Long
                            ' loop through all the data, assuming each row is columns bytes wide
                            ' note: each row has columns bytes of data + 1 for PNG filter type, except 1st row
                            columns = columns + 1
                            Dim rowPredictor As Long
                            For rowIndex = LBound(m_udata) To UBound(m_udata)
                                'rowOffset = LBound(m_udata) + ((rowIndex - LBound(m_udata)) Mod columns)
                                rowOffset = rowIndex Mod columns ' simplify since LBound(m_udata) = 0
                                ' get this rows predicator tag, Note: 1st row doesn't have tag byte
                                If rowOffset = 0 Then ' this byte is a tag byte for new row
                                    rowPredictor = m_udata(rowIndex)
                                    ' advance to actual data
                                    rowIndex = rowIndex + 1
                                    If rowIndex > UBound(m_udata) Then Exit For
                                End If
                                                        
                                Select Case rowPredictor
                                    Case 0, 10 ' PNG, None on all rows
                                        buffer(bufferIndex) = m_udata(rowIndex)
                                    Case 1, 11 ' PNG, Sub on all rows
                                        Stop ' no implemented
                                    Case 2, 12 ' PNG, Up on all rows
                                        If rowIndex >= columns Then ' assume 1st row, with prior row values always 0, so no change to values
                                            ' assume buffer(bufferIndex - (columns-1)) is the data byte 1 row Up, add that to current difference value at rowIndex
                                            buffer(bufferIndex) = (CLng(buffer(bufferIndex - (columns - 1))) + CLng(m_udata(rowIndex))) Mod 256
                                            'Debug.Print Hex(buffer(bufferIndex))
                                        Else
                                            buffer(bufferIndex) = m_udata(rowIndex)
                                        End If
                                    Case 3, 13 ' PNG, Average on all rows
                                        Stop ' no implemented
                                    Case 4, 14 ' PNG, Paeth on all rows
                                        Stop ' no implemented
                                    Case 15 ' PNG, Optimal (per row determination)
                                        Stop ' no implemented, error not valid PNG filter value
                                            
                                    Case Else
                                        Debug.Print "Error: unsupported or invalid predictor found - " & predictor
                                        Stop
                                End Select ' rowPredictor
                                
                                bufferIndex = bufferIndex + 1
                                DoEvents
                            Next rowIndex
                            ' swap out with our smaller buffer omitting predicator data
                            m_udata = buffer
                            ReDim Preserve m_udata(0 To (bufferIndex - 1))
                            
                        Case Else
                            Debug.Print "Error: unsupported or invalid predictor found - " & predictor
                            Stop
                    End Select ' predictor
                Case "/lzwdecode"
                    Stop ' TODO
                Case Else
                    Stop ' not yet supported!
            End Select
        Else
            m_udata = data
        End If
        
        decompressed = True
        udata = m_udata
    End If
End Property

' returns meta data as a VBA Dictionary object
Public Property Get Meta() As Dictionary
    Set Meta = stream_meta.value
End Property

Public Sub Init(ByRef Meta As pdfValue, ByRef data As pdfValue)
    Set stream_meta = Meta
    Set stream_data = data
    decompressed = False
End Sub

Public Sub Class_Initialize()
    Set stream_meta = New pdfValue
    stream_meta.valueType = PDF_ValueType.PDF_Null
End Sub

Public Sub Class_Terminate()
    Set stream_meta = Nothing
    Set stream_data = Nothing
    Erase m_udata
End Sub
