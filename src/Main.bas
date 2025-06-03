Attribute VB_Name = "Main"
Option Explicit


' reads in PDF description and all obj in catalog (does not read in sequentially, nor any unreferenced obj)
Sub LoadPDF(ByRef pdfFilename As String, _
            ByRef trailer As pdfValue, _
            ByRef xrefTable As Dictionary, _
            ByRef info As pdfValue, _
            ByRef root As pdfValue, _
            ByRef pdfObjs As Dictionary)
    On Error GoTo errHandler
    If pdfObjs Is Nothing Then Set pdfObjs = New Dictionary
    
    Dim content() As Byte
    Dim fileLen As Long
    content = readFile(pdfFilename, fileLen)
    If fileLen < 1 Then
        MsgBox "Error reading in pdf", vbOKOnly Or vbCritical, pdfFilename
        Exit Sub
    End If
    
    ' get start of xref
    'Dim xrefOffset As Long: xrefOffset = GetXrefOffset(content)
    
    ' get trailer with /Root information
    Set trailer = GetTrailer(content)
    
    ' load the xref table
    Set xrefTable = GetXrefTable(content, trailer)
    
    ' display info
    Set info = GetInfoObject(content, trailer, xrefTable)
    Debug.Print BytesToString(serialize(info))
    
    ' root obj of PDF
    Set root = GetRootObject(content, trailer, xrefTable)
    Debug.Print BytesToString(serialize(root))
    
    pdfObjs.Add root.id, root
    'pdfObjs.Add info.id info
    GetObjectsInTree root, content, xrefTable, pdfObjs
    
    Exit Sub
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Sub


' writes string to file as bytes and returns count of bytes written
Function PutString(ByVal fileNum As Integer, ByRef str As String) As Long
    Dim data() As Byte: data = StringToBytes(str)
    PutString = PutBytes(fileNum, data)
End Function


' writes bytes to file and returns count of bytes written
Function PutBytes(ByVal fileNum As Integer, ByRef data() As Byte) As Long
    On Error GoTo errHandler
    Dim byteCount As Long
    byteCount = UBound(data) - LBound(data) + 1
    Put #fileNum, , data
    PutBytes = byteCount
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function


' returns a catalog with required object id 0 free
Function NewxrefTable() As Dictionary
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
    Set NewxrefTable = xrefTable
    Set xrefTable = Nothing
End Function


' updates offset or adds entry to catalog for obj
Sub AddUpdateXref(ByRef obj As pdfValue, ByVal offset As Long, ByRef xrefTable As Dictionary, Optional ByVal baseId As Long = 0)
    On Error GoTo errHandler
    Dim entry As xrefEntry
    Dim id As Long
    id = baseId + obj.id
    If xrefTable.Exists(id) Then
        If IsEmpty(xrefTable.Item(id)) Then
            Stop
            GoTo newEntry
        End If
        ' update existing entry
        Set entry = xrefTable.Item(id)
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



' SavePDF split into 3 parts
Function SavePdfHeader(ByRef pdfFilename As String, ByRef offset As Long, Optional ByVal header As String) As Integer
    ' delete if file exists, as otherwise may be extra junk at end of file, but ignore if doesn't exist or other error
    On Error Resume Next
    Kill pdfFilename
    On Error GoTo errHandler
    
    Dim outputFileNum As Integer
    outputFileNum = FreeFile
    Open pdfFilename For Binary Access Write Lock Write As #outputFileNum
    
    
    Const defHeader As String = "%PDF-1.7" & vbNewLine
    If IsBlank(header) Then header = defHeader
    offset = PutString(outputFileNum, header)

    SavePdfHeader = outputFileNum
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Function

Sub SavePdfObjects(ByRef outputFileNum As Integer, _
            ByRef xrefTable As Dictionary, _
            ByRef pdfObjs As Dictionary, _
            ByRef offset As Long, _
            Optional ByVal baseId As Long = 0)
    On Error GoTo errHandler
    
    'SaveObjs()
    Dim v As Variant
    For Each v In pdfObjs.Items
        Dim obj As pdfValue
        Set obj = v
        offset = offset + PutString(outputFileNum, vbLf)
        AddUpdateXref obj, offset, xrefTable, baseId
        offset = offset + PutBytes(outputFileNum, serialize(obj, baseId))
    Next v
    
    Exit Sub
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Sub

Sub SavePdfTrailer(ByRef outputFileNum As Integer, _
            ByRef trailer As pdfValue, _
            ByRef xrefTable As Dictionary, _
            ByRef info As pdfValue, _
            ByRef root As pdfValue, _
            ByRef offset As Long, _
            Optional ByVal prettyPrint As Boolean = True)
    On Error GoTo errHandler

    If prettyPrint Then offset = offset + PutString(outputFileNum, vbLf)
    AddUpdateXref info, offset, xrefTable
    offset = offset + PutBytes(outputFileNum, serialize(info, 0))
    If prettyPrint Then offset = offset + PutString(outputFileNum, vbLf)
    
    ' output xref catalog, for simple form, order should match id#s
    ' each entry should be exactly 20 bytes include 2 character whitespace so needs to end with \r\n or <space>\r or <space>\n
    ' Note: we may leave some id's unused, so we need to actually calculate our highest id and cycle through that in order
    Dim v As Variant
    Dim entry As xrefEntry
    Dim maxId As Long
    maxId = xrefTable.Count - 1 ' should be at least this high
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
            Set entry = xrefTable.Item(ndx)
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
    Dim obj As pdfValue
    Set obj = trailer.Value.Value.Item("/Size")
    obj.Value = CLng(maxId + 1) ' pdfValueObj(CLng(maxId))
    PutBytes outputFileNum, serialize(trailer, 0)

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
            


' given a set of PDF obj and initialized xref (read in or from NewXref), writes out a copy of the PDF
Sub SavePdf(ByRef pdfFilename As String, _
            ByRef trailer As pdfValue, _
            ByRef xrefTable As Dictionary, _
            ByRef info As pdfValue, _
            ByRef root As pdfValue, _
            ByRef pdfObjs As Dictionary)
    On Error GoTo errHandler
    
    Dim offset As Long
    Dim outputFileNum As Integer
    outputFileNum = SavePdfHeader(pdfFilename, offset)
    
    SavePdfObjects outputFileNum, xrefTable, pdfObjs, offset
    SavePdfTrailer outputFileNum, trailer, xrefTable, info, root, offset
    
    Debug.Print "Saved " & pdfFilename
    Exit Sub
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
End Sub


' find /Pages reference in /Root/Catalog and returns corresponding object from pdfObjs Dictionary
Private Function FindPages(ByRef rootCatalog As pdfValue, ByRef pdfObjs As Dictionary) As pdfValue
    On Error GoTo errHandler
    
    Dim obj As pdfValue
    Set obj = rootCatalog.Value ' the dictionary << >>
    Dim dict As Dictionary
    Set dict = obj.Value
    If dict.Exists("/Pages") Then
        Set obj = dict.Item("/Pages")
        ' obj should now be a reference to our /Pages object
        If obj.valueType = PDF_ValueType.PDF_Reference Then
            Dim id As Long, generation As Long
            id = obj.Value
            generation = obj.generation
            Dim v As Variant
            For Each v In pdfObjs.Items
                Set obj = v
                If obj.id = id And obj.generation = generation Then
                    Set FindPages = obj
                    Exit For
                End If
            Next v
        ElseIf obj.valueType = PDF_ValueType.PDF_Object Then
            ' ok, weird but whatever
            Set FindPages = obj
        Else
            ' error! didn't find what we expected here
            Stop
        End If
    End If
    
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Function


' given two /Pages objects, combines them and returns unified /Pages object
' Note: /Pages objects from 2nd pages are bias'd so id begins at baseId+1 (where obj with id=0 is always assumed to be root of free list)
' Warning: this will modify any VBA Objects for pdfValue Reference objects contained in pages2 /Kids
Private Function CombinePages(ByRef pages1 As pdfValue, ByRef pages2 As pdfValue, Optional ByVal baseId As Long) As pdfValue
    Dim obj As pdfValue
    Dim pages As pdfValue
    Set pages = New pdfValue
    pages.id = pages1.id
    pages.generation = 0
    pages.valueType = PDF_ValueType.PDF_Object
    Dim dict As Dictionary: Set dict = New Dictionary
    Set obj = New pdfValue
    obj.valueType = PDF_ValueType.PDF_Dictionary
    Set obj.Value = dict
    Set pages.Value = obj
    
    Dim kids As Collection: Set kids = New Collection
    
    Dim v As Variant
    Dim d As Dictionary
    Dim c As Collection
    
    Set d = pages1.Value.Value
    Set obj = d.Item("/Kids")
    Set c = obj.Value
    For Each v In c
        Set obj = v ' obj reference
        kids.Add obj
    Next v
    
    ' Warning! TODO! need to fixup actual obj Referenced in here so its /Parent refers back to our pages.id obj
    Set d = pages2.Value.Value
    Set obj = d.Item("/Kids")
    Set c = obj.Value
    For Each v In c
        Set obj = v ' obj reference
        obj.Value = baseId + obj.Value      ' Note: only done for pages2 /Kids, this is a reference, so value is the id we reference
        kids.Add obj
    Next v
    
    ' create our new /Count and /Kids objs
    dict.Add "/Type", pdfNameObj("/Pages")
    dict.Add "/Count", pdfValueObj(kids.Count)
    dict.Add "/Kids", pdfArrayObj(kids)
    
    Set CombinePages = pages
    Set pages = Nothing
End Function


Public Sub CombinePDFs(ByRef sourceFiles() As String, ByRef outFile As String)
    Dim oTrailer As pdfValue
    Dim oXrefTable As Dictionary: Set oXrefTable = NewxrefTable()
    Dim oInfo As pdfValue
    Dim oRoot As pdfValue
    Dim oPages As pdfValue  ' /Type /Pages with /Count # /Kids [ /Page references ...]
    'Dim oPdfObjs as Dictionary: Set oPdfObjs = New Dictionary
    
    Dim offset As Long
    Dim outputFileNum As Integer
    outputFileNum = SavePdfHeader(outFile, offset)
    
    Dim ndx As Long, baseId As Long
    For ndx = LBound(sourceFiles) To UBound(sourceFiles)
        Dim trailer As pdfValue: Set trailer = Nothing
        Dim xrefTable As Dictionary: Set xrefTable = Nothing
        Dim info As pdfValue: Set info = Nothing
        Dim root As pdfValue: Set root = Nothing
        Dim pages As pdfValue: Set pages = Nothing
        Dim pdfObjs As Dictionary: Set pdfObjs = Nothing
        
        Debug.Print "Loading file " & ndx + 1 & " - " & sourceFiles(ndx)
        Application.StatusBar = ndx + 1 & " - " & sourceFiles(ndx)
        LoadPDF sourceFiles(ndx), trailer, xrefTable, info, root, pdfObjs
        
        ' for each additional document we need to update /Pages
        Set pages = FindPages(root, pdfObjs)
        
        If oRoot Is Nothing Then
            ' first time through we can just use /Root form pdf unchanged
            Set oRoot = root
            Set oTrailer = trailer
            Set oInfo = info
            Set oPages = pages
            
            ' so remove pages object for now, we add a single one back after we have gone through all of them
            pdfObjs.Remove pages.id
        Else
            ' each additional pdf need to remove /Root, so only left with 1 (the first) /Root
            pdfObjs.Remove root.id
#If False Then
            ' so remove pages object for now, we add a single one back after we have gone through all of them
            pdfObjs.Remove pages.id
            Set oPages = CombinePages(oPages, pages, baseId)
#Else
            ' instead of removing /Pages and combining into a single /Pages, we add a higher level /Pages
            ' that poitns to all /Pages, but we need to fixup their parents to point to our new /Pages
            ' so we use 1st /Pages as new parent for all future /Pages
            Dim obj As pdfValue
            Set obj = New pdfValue
            obj.valueType = PDF_ValueType.PDF_Reference
            obj.Value = oPages.id - baseId  ' Warning! this may be negative, but corrects when Saved and baseId added back!
            If pages.Value.Value.Exists("/Parent") Then
                ' we don't handle this case, it shouldn't exist
                Stop
            Else
                Set pages.Value.Value.Item("/Parent") = obj ' should be Reference obj
            End If
            ' and add to our top level /Pages
            Set obj = oPages.Value.Value.Item("/Kids")
            Dim c As Collection
            Set c = obj.Value
            ' but as a reference
            Set obj = New pdfValue
            obj.valueType = PDF_ValueType.PDF_Reference
            obj.Value = pages.id + baseId   ' we must correct id here as we don't when saving this obj
            c.Add obj
            ' we also need to update our count
            Set obj = oPages.Value.Value.Item("/Count")
            ' ### not +1 as not a count of /Kids but /Count of total pages, so need to sum child counts
            'obj.value = CLng(obj.value + 1)
            Dim Count As Long
            Count = CLng(pages.Value.Value.Item("/Count").Value)
            obj.Value = CLng(obj.Value + Count)
#End If
        End If
        
        SavePdfObjects outputFileNum, oXrefTable, pdfObjs, offset, baseId
        
        ' determine highest id used, 1st obj in next file will start at this + 1
        baseId = baseId + xrefTable.Count - 1 ' highest id possible so far
        DoEvents
    Next ndx
    ' replace the /Pages object
    'Set oPdfObjs(oPages.id) = oPages
    
    ' save updated /Pages
    offset = offset + PutString(outputFileNum, vbLf)
    AddUpdateXref oPages, offset, oXrefTable
    offset = offset + PutBytes(outputFileNum, serialize(oPages, 0))
    
    SavePdfTrailer outputFileNum, oTrailer, oXrefTable, oInfo, oRoot, offset
    'SavePdf outFile, trailer, oXrefTable, info, root, oPdfObjs
    Debug.Print "Saved " & outFile
End Sub


' simple function lets user pick files to combine
Public Sub PickAndCombinePdfFiles()
    Dim files() As String
    files = PickFiles()
    Dim ufFileOrder As ufFileList: Set ufFileOrder = New ufFileList
    ufFileOrder.list = files
    ufFileOrder.Show
    files = ufFileOrder.list
    CombinePDFs files, "combined.pdf"
End Sub

Private Function PickFiles() As String()
    'Create a FileDialog object as a File Picker dialog box.
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = True
        .Title = "Select PDF files to combine:"
        .InitialFileName = "C:\Users\jeremyd\Downloads"
        .Filters.Clear
        .Filters.Add "PDF files", "*.pdf"
        .Filters.Add "All files", "*.*"
        '.FilterIndex = 1
 
        'Use the Show method to display the File Picker dialog box and return the user's action.
        'The user pressed the button.
        If .Show = -1 Then
            'Step through each string in the FileDialogSelectedItems collection.
            Dim files() As String
            ReDim files(0 To .SelectedItems.Count - 1)
            Dim ndx As Long
            For ndx = 1 To .SelectedItems.Count
                files(ndx - 1) = .SelectedItems(ndx)
            Next ndx
            PickFiles = files
        Else 'The user pressed Cancel.
        End If
        
        .Filters.Clear
    End With
 
    Set fd = Nothing
End Function
