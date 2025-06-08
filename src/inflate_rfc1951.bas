Attribute VB_Name = "inflate_rfc1951"
' inverse of RFC1951 Deflate, inflates (decompresses) data compressed with the deflate algorithm
' Note: allows raw deflate or RFC1950 zlib wrapped deflate data, does not handle RFC1952 gzip wrapper
Option Explicit


Private Enum ZlibCompressionType
    NoCompression = 0
    FixedHuffmanCodes = 1
    DynamicHuffmanCodes = 2
    Reserved_Error = 3
End Enum


'Private Type huffmanCode
'    bitLen As Integer
'    code As Long
'End Type

' LiteralCode has values 0 to 255, with 256 end block and > 256 to 285 are length codes
Private literalAndLenCodes() As huffmanCode
Private mapLiteralAndLenHuffmanCodeToValue() As Integer ' index corresponds with huffman code and returns corresponding value for said code
' Note mapLiteralAndLenHuffmanCodeToValue(literalAndLenCodes(N).code) == N
' i.e. given value N, the huffman code H = literalAndLenCodes(N).code
'  and given huffman code H, the value N = mapLiteralAndLenHuffmanCodeToValue(H)

'Private Type codeNode
'    extraBitsLen As Integer    ' how many extra bits to extract and add to min value
'    minValue As Long           ' starting value for the corresponding code this node represents
'    maxValue As Long           ' max value for the code value this node represents
'End Type

Private distanceTree(0 To 31) As codeNode
Private lengthTree(0 To (285 - 257)) As codeNode    ' (0 to 28)

' given the match distance code (value of 1 to 31) returns how many extra bits needed to get full distance value
' match distance table consists of 32 symbols followed by extra bits to encode distance in output window to start of previously seen value
' actual distance is in range of 1 to 32768, where value is based on 5 bit code + N extra bits where N = Floor(N/2)-1 for N>1
Private Function distanceExtraBits(ByVal code As Integer) As Integer
    If code < 4 Then    ' handles case of N <= 1
        distanceExtraBits = 0
    Else
        distanceExtraBits = Int(code / 2) - 1
    End If
End Function

' given the length code (value 0 to 28) returns how many extra bits needed to get full length value
' Note: length is encoded as values 257 to 285, we expect code to be range 0 to 28 (257-257 to 285-257)
' which represent length values from 3 to 258 when including extra bits
Private Function lengthExtraBits(ByVal code As Integer) As Integer
    If code < 8 Or code > 27 Then
        lengthExtraBits = 0
    Else
        lengthExtraBits = (code - 4) \ 4
    End If
End Function


' given distance code (1 to 31) and Byte buffer along with Byte & Bit index into buffer to read next bits from,
' Returns full distance value, additionally advances byteOffset and bitOffset correspondingly (by number of extra bits read)
Private Function distanceValue(ByVal code As Integer, ByRef bytes() As Byte, ByRef byteOffset As Long, ByRef bitOffset) As Long
    If code < 0 Or code > 31 Then GoTo badCode
    With distanceTree(code)
        Dim extraBits As Long
        extraBits = nBITS(.extraBitsLen, bytes, byteOffset, bitOffset)
        distanceValue = .minValue + extraBits
        If distanceValue > .maxValue Then GoTo badCode
    End With
    Exit Function
badCode:
    Err.Raise 9, "Deflate", "Invalid distance code for " & code
End Function

' given length code (0 to 28) and Byte buffer along with Byte & Bit index into buffer to read next bits from,
' Returns full length value from 3 to 358, additionally advances byteOffset and bitOffset correspondingly (by number of extra bits read)
Private Function lengthValue(ByVal code As Integer, ByRef bytes() As Byte, ByRef byteOffset As Long, ByRef bitOffset) As Long
    If code < 0 Or code > 28 Then GoTo badCode
    With lengthTree(code)
        Dim extraBits As Long
        extraBits = nBITS(.extraBitsLen, bytes, byteOffset, bitOffset)
        lengthValue = .minValue + extraBits
        If lengthValue > .maxValue Then GoTo badCode
    End With
    Exit Function
badCode:
    Err.Raise 9, "Deflate", "Invalid length code for " & code
End Function

Private Sub buildDistanceTree()
    Dim bytes(0 To 3) As Byte
    bytes(0) = 255: bytes(1) = 255: bytes(2) = 255: bytes(3) = 255  ' max code's distance has all 1s in its extra bits
    Dim currentMax As Long
    Dim distanceCode As Long ' first 5 bits are code value for distance, 0-31 + extra bits
    For distanceCode = 0 To 31
        Set distanceTree(distanceCode) = New codeNode
        distanceTree(distanceCode).extraBitsLen = distanceExtraBits(distanceCode)
        
        Dim byteOffset As Long, bitOffset As Long
        byteOffset = 0: bitOffset = 0
        Dim extraBits As Long
        extraBits = nBITS(distanceTree(distanceCode).extraBitsLen, bytes, byteOffset, bitOffset)
        distanceTree(distanceCode).minValue = currentMax + 1 '+ 0          ' this code's distance begins 1 greater than last code's max value
        distanceTree(distanceCode).maxValue = currentMax + 1 + extraBits   ' max distance value used is 32768
        currentMax = distanceTree(distanceCode).maxValue                   ' update for next iteration of loop
        
        Debug.Print "Code=" & distanceCode & " (" & distanceTree(distanceCode).extraBitsLen & " bits) " & distanceTree(distanceCode).minValue & "-" & distanceTree(distanceCode).maxValue
    Next
End Sub

Private Sub buildLengthTree()
    Dim bytes(0 To 3) As Byte
    bytes(0) = 255: bytes(1) = 255: bytes(2) = 255: bytes(3) = 255  ' max code's distance has all 1s in its extra bits
    Dim currentMax As Long: currentMax = 2
    Dim lengthCode As Long ' encoded value - 257 so in range of 0 to 28
    For lengthCode = 0 To 27
        Set lengthTree(lengthCode) = New codeNode
        lengthTree(lengthCode).extraBitsLen = lengthExtraBits(lengthCode)
        
        Dim byteOffset As Long, bitOffset As Long
        byteOffset = 0: bitOffset = 0
        Dim extraBits As Long
        extraBits = nBITS(lengthTree(lengthCode).extraBitsLen, bytes, byteOffset, bitOffset)
        lengthTree(lengthCode).minValue = currentMax + 1 '+ 0          ' this code's distance begins 1 greater than last code's max value
        lengthTree(lengthCode).maxValue = currentMax + 1 + extraBits
        currentMax = lengthTree(lengthCode).maxValue               ' update for next iteration of loop
        
        Debug.Print "Code=" & lengthCode + 257 & " (" & lengthTree(lengthCode).extraBitsLen & " bits) " & lengthTree(lengthCode).minValue & "-" & lengthTree(lengthCode).maxValue
    Next
    
    ' handle last code specially
    Set lengthTree(28) = New codeNode
    With lengthTree(28)
        .extraBitsLen = 0
        .minValue = 258
        .maxValue = 258
    End With
End Sub


Private Sub showDistanceTree()
    Dim bytes(0 To 3) As Byte
    Dim byteOffset As Long, bitOffset As Long
    buildDistanceTree
    
    Dim code As Long
    For code = 0 To 31
        byteOffset = 0: bitOffset = 0
        Debug.Print "Code " & code & " starts at " & distanceValue(code, bytes, byteOffset, bitOffset)
        Debug.Print "Used " & byteOffset & " bytes with " & bitOffset & " bits"
    Next code
End Sub

Private Sub showLengthTree()
    Dim bytes(0 To 3) As Byte
    Dim byteOffset As Long, bitOffset As Long
    buildLengthTree
    
    Dim code As Long
    For code = 0 To 28
        byteOffset = 0: bitOffset = 0
        Debug.Print "Code " & code + 257 & " starts at " & lengthValue(code, bytes, byteOffset, bitOffset)
        'Debug.Print "Used " & byteOffset & " bytes with " & bitOffset & " bits"
    Next code
End Sub


' given an array of numerical values, the index corresponding with the code and the value the X/--frequency (count) of that code--/X
' given an array of numerical values, the index corresponding with the code and the value the length of that code, shorter being most frequent
' returns an array of huffman coded values
Private Function buildHuffmanTree(ByRef lengths As Variant, ByRef minBitLen As Integer, ByRef maxBitLen As Integer) As huffmanCode()
    Dim curCode As Long
    Dim curBitLength As Integer
    Dim minCode As Long, maxCode As Long
    minCode = LBound(lengths): maxCode = UBound(lengths)
    Dim huffmanTree() As huffmanCode
    ReDim huffmanTree(minCode To maxCode)   ' initialized so all values initially 0
    
    ' find smallest and largest values, 7 to 9 for default static Huffman tree
    Dim minValue As Long, maxValue As Long
    minValue = &H7FFFFFFF: maxValue = -1
    Dim ndx As Long
    For ndx = minCode To maxCode
        Dim Value As Long: Value = CLng(lengths(ndx))
        If (Value < minValue) And (Value > 0) Then
            minValue = Value
        End If
        If Value > maxValue Then
            maxValue = Value
        End If
        DoEvents
    Next ndx
    minBitLen = minValue
    maxBitLen = maxValue
    
    ' we assume first code is all 0 bits
    curCode = 0
    ' we look for lexiconally most frequent code index (shortest length) and assign its code as right side of tree, the proceed down a level (adding a 1 bit + ? more bits)
    ' loop through all length values updating our tree correspondingly
    For curBitLength = minValue To maxValue
        ' we need to left shift the curCode once for each additional bit (Note: initial value is 0 so first time through does nothing)
        curCode = LShift(curCode, 1)
        For ndx = minCode To maxCode
            'If CLng(freq(ndx)) = curFreq Then
            If (CLng(lengths(ndx)) = curBitLength) And (curBitLength > 0) Then
                ' right side of tree
                'Debug.Print ndx & "=" & Hex(curCode)
                Set huffmanTree(ndx) = New huffmanCode
                With huffmanTree(ndx)
                    .bitLen = curBitLength
                    .code = curCode
                End With
                ' left side of tree
                curCode = curCode + 1
            End If
            DoEvents
        Next ndx
    Next curBitLength
    
    buildHuffmanTree = huffmanTree
End Function

' literals and length values, 0 to 287
Private Function buildStaticHuffmanTree() As huffmanCode()
    Dim bitLengths(0 To 287) As Byte
    Dim ndx As Long
    For ndx = LBound(bitLengths) To UBound(bitLengths)
        Select Case ndx
            Case 0 To 143 ' codes   0011,0000 -   1011,1111
                bitLengths(ndx) = 8
            Case 144 To 255 ' codes 1,1001,0000 - 1,1111,1111
                bitLengths(ndx) = 9
            Case 256 To 279 ' codes    000,0000 -    001,0111
                bitLengths(ndx) = 7
            Case 280 To 287 ' codes   1100,0000 -   1100,0111
                bitLengths(ndx) = 8
            Case Else
                Stop    ' error!
        End Select
    Next ndx
    
    Dim minBitLen As Integer, maxBitLen As Integer
    buildStaticHuffmanTree = buildHuffmanTree(bitLengths, minBitLen, maxBitLen)
    If (minBitLen <> 7) Or (maxBitLen <> 9) Then
        Debug.Print "Bit lengths for static huffman tree unexpected length!"
        Stop
    End If
End Function

' generate reverse mapping, assumes largest huffman code is maxBits long
Private Sub buildReverseHuffmanMapping(ByRef mapHuffmanCodeToValue() As Integer, ByRef huffmanCodes() As huffmanCode, ByVal maxBits As Integer)
    Dim maxCode As Long
    maxCode = LShift(1, maxBits) - 1
    ReDim mapHuffmanCodeToValue(0 To maxCode)
    
    ' loop through all values and using its huffman code as index into our mapping table
    ' set the value to our index (the reverse mapping)
    ' Note: unused huffman codes will map to the value 0, so we explicitly initialize all values to -1 to detect invalid codes
    Dim ndx As Long
    For ndx = LBound(mapHuffmanCodeToValue) To UBound(mapHuffmanCodeToValue)
        mapHuffmanCodeToValue(ndx) = -1
    Next ndx
    For ndx = LBound(huffmanCodes) To UBound(huffmanCodes)
        If Not huffmanCodes(ndx) Is Nothing Then
            With huffmanCodes(ndx)
                'Debug.Print ndx & "(" & .bitLen & ") " & Hex(.code) & " [" & .code & "]"
                mapHuffmanCodeToValue(.code) = ndx
            End With
        End If
    Next ndx
End Sub


' initialize all our mapping tables
Sub initializeStaticCodeTables()
    literalAndLenCodes = buildStaticHuffmanTree
    ' our codes are all 7,8, or 9 bits
    buildReverseHuffmanMapping mapLiteralAndLenHuffmanCodeToValue, literalAndLenCodes, 9
End Sub

' initialize dynamic tables
' input is an array of bitLengths used to build huffman table
' output is array of huffmanCodes, array mapping huffman code to value, and min/maxBitLen
Sub intializeDynamicCodeTables(ByRef bitLengths() As Byte, ByRef mapHuffmanCodeToValue() As Integer, _
                                      ByRef huffmanCodes() As huffmanCode, _
                                      ByRef minBitLen As Integer, ByRef maxBitLen As Integer)
    ' build the huffman tree
    huffmanCodes = buildHuffmanTree(bitLengths, minBitLen, maxBitLen)
    ' generate reverse mapping so we can take convert huffman code to original value
    buildReverseHuffmanMapping mapHuffmanCodeToValue, huffmanCodes, maxBitLen
End Sub

' in: cbytes, byteOffset, bitOffset, minBits, maxBits, mapHuffmanCodeToValue, huffmanCodes, countCodeValues
' out: bitLengths, updated byteOffset and bitOffset
Sub readInDynamicHuffmanBitLengths(ByRef cbytes() As Byte, ByRef byteOffset As Long, ByRef bitOffset As Long, _
                                  ByVal minBits As Integer, ByVal maxBits As Integer, _
                                  ByRef mapHuffmanCodeToValue() As Integer, ByRef huffmanCodes() As huffmanCode, _
                                  ByRef bitLengths() As Byte, ByVal countCodeValues As Integer)
    ReDim bitLengths(0 To countCodeValues - 1)
    Dim i As Integer
    For i = 0 To countCodeValues - 1
        Dim codedLength As Integer
        Dim repeatCount As Integer
        codedLength = readHuffmanValue(cbytes, byteOffset, bitOffset, minBits, maxBits, mapHuffmanCodeToValue, huffmanCodes)
    
        ' 0 - 15: Represent code lengths of 0 - 15
        ' 16: Copy the previous code length 3 - 6 times.
        '       The next 2 bits indicate repeat length
        '          (0 = 3, ... , 3 = 6)
        '          Example:  Codes 8, 16 (+2 bits 11),
        '                 16 (+2 bits 10) will expand to
        '                 12 code lengths of 8 (1 + 6 + 5)
        ' 17: Repeat a code length of 0 for 3 - 10 times.
        '     (3 bits of length)
        ' 18: Repeat a code length of 0 for 11 - 138 times
        '     (7 bits of length)
        
        Select Case codedLength
            Case 0 To 15:
                'bitLengths(i) = codedLength
                repeatCount = 1
            Case 16:
                repeatCount = nBITS(2, cbytes, byteOffset, bitOffset) + 3
                codedLength = bitLengths(i - 1) ' previous code length
            Case 17:
                repeatCount = nBITS(3, cbytes, byteOffset, bitOffset) + 3
                codedLength = 0
            Case 18:
                repeatCount = nBITS(7, cbytes, byteOffset, bitOffset) + 11
                codedLength = 0
            Case Else
                Stop
        End Select
        Dim j
        For j = i To i + repeatCount - 1
            bitLengths(j) = codedLength
        Next j
        i = i + repeatCount - 1
    Next i
End Sub

Function readInDynamicHuffmanTable(ByRef cbytes() As Byte, ByRef byteOffset As Long, ByRef bitOffset As Long, _
    ByRef mapLiteralLengthHuffmanCodeToValue() As Integer, ByRef literalLengthHuffmanCodes() As huffmanCode, _
    ByRef literalLengthMinBits As Integer, ByRef literalLengthMaxBits As Integer, _
    ByRef mapDistanceHuffmanCodeToValue() As Integer, ByRef distanceHuffmanCodes() As huffmanCode, _
    ByRef distanceMinBits As Integer, ByRef distanceMaxBits As Integer _
) As Boolean
    On Error GoTo errHandler
    '5 Bits: HLIT, # of Literal/Length codes - 257 (257 - 286)
    '5 Bits: HDIST, # of Distance codes - 1        (1 - 32)
    '4 Bits: HCLEN, # of Code Length codes - 4     (4 - 19)
    Dim countLiteralCodeValue As Integer  ' HLIT+257
    Dim countDistanceCodeValue As Integer ' HDIST+1
    Dim countCodes As Integer           ' HCLEN+4
    countLiteralCodeValue = nBITS(5, cbytes, byteOffset, bitOffset) + 257
    countDistanceCodeValue = nBITS(5, cbytes, byteOffset, bitOffset) + 1
    countCodes = nBITS(4, cbytes, byteOffset, bitOffset) + 4

    ' read in bit counts for huffman code our dynamic huffman code lengths encoded in
    Dim order() As Variant
    order = Array(16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15)
    Dim codeLens() As Byte
    ReDim codeLens(0 To 18) ' if countCodes < 18, indices at end of order array have codeLens(i) = 0
    Dim i As Integer
    For i = 0 To countCodes - 1
        codeLens(order(i)) = nBITS(3, cbytes, byteOffset, bitOffset)
    Next i
    ' generate huffman mapping that dynamic bit lengths are encoded in
    Dim mapCodeLenHuffmanCodeToValue() As Integer
    Dim codeLenHuffmanCodes() As huffmanCode
    Dim codeLenMinBits As Integer, codeLenMaxBits As Integer
    intializeDynamicCodeTables codeLens, mapCodeLenHuffmanCodeToValue, codeLenHuffmanCodes, codeLenMinBits, codeLenMaxBits
    
    Dim bitLengths() As Byte
    
    ' read in dynamic literal/length bit lengths (encoded using codeLen generated huffman tree)
    readInDynamicHuffmanBitLengths cbytes, byteOffset, bitOffset, codeLenMinBits, codeLenMaxBits, mapCodeLenHuffmanCodeToValue, codeLenHuffmanCodes, _
            bitLengths, countLiteralCodeValue   ' read in bitLengths array used to determine huffman table mapping
    intializeDynamicCodeTables bitLengths, mapLiteralLengthHuffmanCodeToValue, literalLengthHuffmanCodes, literalLengthMinBits, literalLengthMaxBits
    
    ' read in dynamic literal/length bit lengths (encoded using codeLen generated huffman tree)
    ' Note: 1 distance code of 0 bits indicates only literals are encoded
    readInDynamicHuffmanBitLengths cbytes, byteOffset, bitOffset, codeLenMinBits, codeLenMaxBits, mapCodeLenHuffmanCodeToValue, codeLenHuffmanCodes, _
            bitLengths, countDistanceCodeValue  ' read in bitLengths array used to determine huffman table mapping
    intializeDynamicCodeTables bitLengths, mapDistanceHuffmanCodeToValue, distanceHuffmanCodes, distanceMinBits, distanceMaxBits
    
    readInDynamicHuffmanTable = True
    Exit Function
errHandler:
    Debug.Print Err.Description
    Stop
    Resume
End Function

Private Function readHuffmanValue(ByRef cbytes() As Byte, ByRef inputOffset As Long, ByRef bitOffset As Long, _
                                  ByVal minBits As Integer, ByVal maxBits As Integer, _
                                  ByRef mapHuffmanCodeToValue() As Integer, ByRef huffmanCodes() As huffmanCode)
    ' decode huffman code to value from input stream
    Dim huffmanBits As Integer
    ' the symbol is a huffman code (bit swapped, so msb in bytes lsb & vice versa)
    ' the humman codes are from N to N+1 bits in len, with static table this is 7 to 9 bits
    ' we read in 1st minBits bits and see if valid code, if not continue reading bits until too large or valid code found
    huffmanBits = nBITS(minBits, cbytes, inputOffset, bitOffset)
    ' flip the bits
    huffmanBits = ReverseBits(minBits, huffmanBits)
    ' see if its a valid huffman code
    Dim validCode As Boolean: validCode = False
checkForValidHuffmanCode:
    validCode = (mapHuffmanCodeToValue(huffmanBits) >= 0)
    ' we need to verify that reverse mapping is for correct bits,
    ' e.g. if we only got 7 bits, don't match against code returned by 8 bits
    If validCode Then validCode = (huffmanCodes(mapHuffmanCodeToValue(huffmanBits)).bitLen = minBits)
    If Not validCode Then ' not valid code
        If minBits < maxBits Then
            minBits = minBits + 1
            ' append another bit
            Dim t As Long
            t = nBITS(1, cbytes, inputOffset, bitOffset)
            huffmanBits = LShift(huffmanBits, 1) Or t
            GoTo checkForValidHuffmanCode
        Else
            Debug.Print "Invalid huffman code!"
            Exit Function
        End If
    End If
        
    ' map huffman code we just read in to value it represents
    readHuffmanValue = mapHuffmanCodeToValue(huffmanBits)
End Function

' reads in 3 bit header and sets lastBlock and blockCompressionType accordingly
' format BFINAL(1 bit), BTYPE(2 bits)
Private Function readBlockHeader(ByRef lastBlock As Boolean, ByRef blockComppressionType As ZlibCompressionType, _
                                 ByRef cbytes() As Byte, ByRef inputOffset As Long, ByRef bitOffset As Long) As Boolean
    On Error GoTo errHandler
    
    ' read in BFINAL bit
    Dim BFINAL As Long
    BFINAL = nBITS(1, cbytes, inputOffset, bitOffset)
    lastBlock = (BFINAL <> 0)
    
    ' read in BTYPE bits
    blockComppressionType = nBITS(2, cbytes, inputOffset, bitOffset)
    
    readBlockHeader = True
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function


 Function ReverseBits(ByVal bitCount As Integer, ByVal Value As Long) As Long
    Dim flipped As Long
    Dim ndx As Long
    For ndx = 0 To (bitCount \ 2)
        ' start from highest & lowest swapping, an progressively move in
        Dim highBitIndex As Long, lowBitIndex As Long
        Dim highBitValue As Long, lowBitValue As Long
        Dim shiftDistance As Long
        shiftDistance = bitCount - (2 * ndx) - 1
        highBitIndex = 2 ^ (bitCount - ndx - 1)
        lowBitIndex = 2 ^ ndx
        highBitValue = Value And highBitIndex
        lowBitValue = Value And lowBitIndex
        flipped = flipped Or LShift(lowBitValue, shiftDistance) Or RShift(highBitValue, shiftDistance)
    Next ndx
    ReverseBits = flipped
    
    'Debug.Print Hex(Value And ((2 ^ bitCount) - 1)) & "-" & Hex(ReverseBits)
End Function

Private Function inflateData( _
    ByRef cbytes() As Byte, ByRef oBytes() As Byte, ByRef inputOffset As Long, ByRef outputOffset As Long, ByRef bitOffset As Long, _
    ByRef mapLiteralLengthHuffmanCodeToValue() As Integer, ByRef literalLengthHuffmanCodes() As huffmanCode, _
    ByVal literalLengthMinBits As Integer, ByVal literalLengthMaxBits As Integer, _
    ByRef mapDistanceHuffmanCodeToValue() As Integer, ByRef distanceHuffmanCodes() As huffmanCode, _
    ByVal distanceMinBits As Integer, ByVal distanceMaxBits As Integer, _
    ByRef distanceTree() As codeNode, lengthTree As Variant) As Boolean
'Private literalAndLenCodes() As huffmanCode
'Private mapLiteralAndLenHuffmanCodeToValue() As Integer ' index corresponds with huffman code and returns corresponding value for said code
'Private DistanceTree(0 To 31) As CodeNode
'Private LengthTree(0 To (285 - 257)) As CodeNode    ' (0 to 28)
    On Error GoTo errHandler
    
    Dim endOfBlock As Boolean
    Do ' loop until we reach end of block code, 256
        'If outputOffset >= 69 Then Stop
        ' decode literal/length value from input stream
        ' by mapping huffman code we read in to value it represents
        Dim literalOrLength As Integer
        literalOrLength = readHuffmanValue(cbytes, inputOffset, bitOffset, _
                          literalLengthMinBits, literalLengthMaxBits, mapLiteralLengthHuffmanCodeToValue, literalLengthHuffmanCodes)

        If literalOrLength < 256 Then ' is it a literal value
            ' copy literal byte to output stream
            oBytes(outputOffset) = CByte(literalOrLength)
            outputOffset = outputOffset + 1
        ElseIf literalOrLength = 256 Then ' is it end of block marker
            endOfBlock = True
            ' Note: do not advance to next byte boundary
            ' if needed, it will be done after reading block header, see no compression section
            'Exit Do
        Else ' is it length, distance value, > 256
            ' decode the rest of length from input stream
            literalOrLength = literalOrLength - 257  ' get length as 0 based value instead of beginning at 257
            Dim lengthCode As codeNode
            Set lengthCode = lengthTree(literalOrLength)
            Dim extra As Long
            extra = nBITS(lengthCode.extraBitsLen, cbytes, inputOffset, bitOffset)
            literalOrLength = lengthCode.minValue + extra
            ' decode distance
            Dim distance As Long
            Dim distanceCode As codeNode
            If distanceMinBits = 5 And distanceMaxBits = 5 Then  ' TODO fixed vs dynamic codes
                distance = nBITS(5, cbytes, inputOffset, bitOffset)
                distance = ReverseBits(5, distance)
            Else
                distance = readHuffmanValue(cbytes, inputOffset, bitOffset, _
                          distanceMinBits, distanceMaxBits, mapDistanceHuffmanCodeToValue, distanceHuffmanCodes)
            End If
            Set distanceCode = distanceTree(distance)
            extra = nBITS(distanceCode.extraBitsLen, cbytes, inputOffset, bitOffset)
            distance = distanceCode.minValue + extra
            ' copy length bytes beginning (outputOffset - distance) bytes in oBytes to oBytes
            Dim i As Long
            For i = 0 To literalOrLength - 1
                oBytes(outputOffset + i) = CByte(oBytes(outputOffset - distance + i))
            Next i
            ' update our indices
            outputOffset = outputOffset + literalOrLength
        End If
    'If outputOffset >= 19 Then Stop
    Loop Until endOfBlock
    inflateData = True
    Exit Function
errHandler:
    Debug.Print Err.Description
    Stop
    Resume
End Function


' Given a Deflate compressed input buffer, output buffer, and starting indices, attempts to returned decompressed (Inflated) data
' Note: assumes decoding starts on byte boundary
' cBytes is deflate compressed stream of bytes ' *** TODO also support zlib wrapped data, but not gzip wrapped
' oBytes is output buffer
' inputOffset is byte index to begin decompression at, updated as input bytes are processed
' outputOffset is byte offset to copy uncompressed data to; updated as bytes are decompressed
' Returns True if able to successfully decode data, False on any errors
Public Function inflate2(ByRef cbytes() As Byte, ByRef oBytes() As Byte, ByRef inputOffset As Long, ByRef outputOffset As Long) As Boolean
    
    ' if not allocated, then assume 4X compressed size (may need to increase, possibly up to 16X)
    On Error Resume Next ' if error getting bounds then assume unallocate
    If UBound(oBytes) < LBound(oBytes) Then ReDim oBytes(0 To ((UBound(cbytes) - LBound(cbytes)) * 4) - 1)
    On Error GoTo errHandler
    
''' intialize our constant trees, used by both static and dynamic huffman tree inflation
    buildDistanceTree
    buildLengthTree
    initializeStaticCodeTables
'''
    
    
    ' see rfc1951
    Dim lastBlock As Boolean
    Dim bitOffset As Long
    Do
        Dim blockCompressionType As ZlibCompressionType
        If Not readBlockHeader(lastBlock, blockCompressionType, cbytes, inputOffset, bitOffset) Then
            Debug.Print "Error reading zlib header!"
            Exit Function
        End If
        
        Select Case blockCompressionType
            Case ZlibCompressionType.NoCompression
                ' skip any remaining bits in current partially processed byte, i.e. align reading input to next byte value
                If bitOffset > 0 Then
                    inputOffset = inputOffset + 1 '+ (bitOffset \ 8)
                    bitOffset = 0
                End If
                ' read LEN and NLEN (each is 2 bytes)
                Dim copyLen As Long, negatedCopyLen As Long
                copyLen = nBITS(16, cbytes, inputOffset, bitOffset)
                negatedCopyLen = (Not nBITS(16, cbytes, inputOffset, bitOffset)) And &HFFFF&    ' limit to 16 bit value
                If copyLen <> negatedCopyLen Then
                    Debug.Print "Error invalid uncompressed length value!"
                    Stop
                    Exit Function
                End If
                ' copy LEN bytes of data to output
                CopyBytes cbytes, oBytes, inputOffset, outputOffset, copyLen
                ' update our indices
                inputOffset = inputOffset + copyLen
                outputOffset = outputOffset + copyLen
            Case ZlibCompressionType.DynamicHuffmanCodes
                ' read representation of code trees
                Dim dynLLHuffmanCodeToValue() As Integer
                Dim dynLLHuffmanCodes() As huffmanCode
                Dim LLminBits As Integer, LLmaxBits As Integer
                Dim dynDHuffmanCodeToValue() As Integer
                Dim dynDHuffmanCodes() As huffmanCode
                Dim DminBits As Integer, DmaxBits As Integer
                ' allocate space, read in, and fill out trees
                If Not readInDynamicHuffmanTable(cbytes, inputOffset, bitOffset, _
                    dynLLHuffmanCodeToValue, dynLLHuffmanCodes, LLminBits, LLmaxBits, _
                    dynDHuffmanCodeToValue, dynDHuffmanCodes, DminBits, DmaxBits) Then
                    Debug.Print "Error extracting dynamic huffman codes!"
                    Exit Function
                End If
                ' continue inflating but using dynamic instead of fixed trees
                If Not inflateData(cbytes, oBytes, inputOffset, outputOffset, bitOffset, dynLLHuffmanCodeToValue, dynLLHuffmanCodes, LLminBits, LLmaxBits, dynDHuffmanCodeToValue, dynDHuffmanCodes, DminBits, DmaxBits, distanceTree, lengthTree) Then
                    Debug.Print "Error inflating with dynamic huffman codes!"
                    Exit Function
                End If
            Case ZlibCompressionType.FixedHuffmanCodes
                ' use our precomputed fixed trees
                'initializeStaticCodeTables  ' as these don't change move to earlier initialization phase
                Dim dHuffmanCodeToValue() As Integer
                Dim dHuffmanCodes() As huffmanCode
                If Not inflateData(cbytes, oBytes, inputOffset, outputOffset, bitOffset, mapLiteralAndLenHuffmanCodeToValue, literalAndLenCodes, 7, 9, dHuffmanCodeToValue, dHuffmanCodes, 5, 5, distanceTree, lengthTree) Then
                    Debug.Print "Error inflating with fixed huffman codes!"
                    Exit Function
                End If
            Case Else ' ZlibCompressionType.Reserved_Error
                Debug.Print "Error, invalid block compression type specified!"
                Exit Function
        End Select
        DoEvents
    Loop While Not lastBlock
    
    ReDim Preserve oBytes(0 To outputOffset - 1)
    inflate2 = True
#If False Then ' debug output
    Dim x As Long
    For x = 0 To outputOffset - 1
        Debug.Print Hex(oBytes(x)) & " ";
    Next
    Debug.Print ""
#End If
    Exit Function
errHandler:
    Debug.Print "inflate_rfc1951.inflate() - " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function

