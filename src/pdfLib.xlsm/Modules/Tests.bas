Attribute VB_Name = "Tests"
' tests and scratch area
Option Explicit


' Call the function like this:
Sub TestPdfCombine()
    Const basedir As String = "C:\Users\jeremyd\Downloads\"
    Dim sources() As String
    sources = Split(basedir & "test.pdf," & basedir & "test.pdf", ",")
    CombinePDFs sources, basedir & "Combined.pdf"
End Sub

Sub TestHeaderAndVersion()
    Dim pdfDoc As pdfDocument
    'Set pdfDoc = New pdfDocument
    Set pdfDoc = pdfDocument.pdfDocument
    Debug.Print pdfDoc.version
    Set pdfDoc = New pdfDocument
    Debug.Print "[" & pdfDoc.Header & "]"
End Sub

Sub TestProblemPdfs()
    On Error GoTo errHandler
    Const basedir As String = "C:\Users\jeremyd\Downloads\"
    'Const filename As String = "2025 Request to Expunge Form pdf.pdf"
    Const filename As String = "pdf-association.pdf20examples\pdf20-utf8-test.pdf"
    
    
    ' create VBA object to work with PDF document
    Dim pdfDoc As pdfDocument
    Set pdfDoc = New pdfDocument
    
    ' attempt to load PDF document, initializes trailer and rootCatalog but otherwise does not parse PDF objects contained in document
    If Not pdfDoc.loadPdf(basedir & filename) Then
        Debug.Print "Error loading " & pdfDoc.filename
    End If
    
    ' without parsing whole document just get the metadata about this document
    Debug.Print BytesToString(pdfDoc.Info.serialize)
    Debug.Print BytesToString(pdfDoc.Meta.serialize)
    
    ' actually load all the pdf objects referenced from the root catalog (does not load orphan'd objects)
    If Not pdfDoc.parsePdf() Then
        Debug.Print "Error parsing pdf " & pdfDoc.filename
    End If
    
    Debug.Print pdfDoc.pages.asDictionary.Item("/Count").Value
    
    Dim obj As pdfValue
    Set obj = pdfDoc.getObject(69, 0)
    If obj.valueType = PDF_EndOfDictionary Then
        Debug.Print obj.Value.Value.Length
    End If
    Stop
    
    Exit Sub
errHandler:
    Debug.Print "Error: " & Err.Description
    Stop
    Resume
End Sub

Sub TestReWritePdf()
    Const basedir As String = "C:\Users\jeremyd\Downloads\"
    
    ' create VBA object to work with PDF document
    Dim pdfDoc As pdfDocument
    Set pdfDoc = New pdfDocument
    
    ' attempt to load PDF document, initializes trailer and rootCatalog but otherwise does not parse PDF objects contained in document
    If Not pdfDoc.loadPdf(basedir & "test2.pdf") Then
        Debug.Print "Error loading " & pdfDoc.filename
    End If
    
    ' without parsing whole document just get the metadata about this document
    'Debug.Print BytesToString(pdfDoc.Info.serialize)
    'Debug.Print BytesToString(pdfDoc.Meta.serialize)
    
    ' actually load all the pdf objects referenced from the root catalog (does not load orphan'd objects)
    If Not pdfDoc.parsePdf() Then
        Debug.Print "Error parsing pdf " & pdfDoc.filename
    End If
    
    ' since we've parse pdf, these should just quickly return cached (previously parsed) metadata objects
    Debug.Print pdfDoc.Info.id
    Debug.Print pdfDoc.Meta.id
    
    ' we didn't make any changes, but save as a new file for comparison - we don't support writing object stream objects nor any compression /Filters
    If Not pdfDoc.savePdfAs(basedir & "rewritten.pdf") Then
        Debug.Print "Error saving " & "rewritten.pdf"
    End If
End Sub

Sub TestZip()
    Const testFile As String = "[Content_Types].xml"
    Dim zip As ExcelZIP: Set zip = New ExcelZIP
    zip.Init ThisWorkbook.Path & "\\" & ThisWorkbook.name
    Dim fileList() As String
    fileList = zip.GetFileNames()
    Dim i As Long
    For i = LBound(fileList) To UBound(fileList)
        Debug.Print fileList(i)
    Next i
    If zip.HasFile(testFile) Then
        Dim fileContent() As Byte
        On Error Resume Next
        zip.ReadData testFile, fileContent
        If Err.Number <> 0 Then
            Debug.Print Err.Description & " (" & Err.Number & ")"
        Else
            For i = LBound(fileContent) To UBound(fileContent)
                Debug.Print Chr(fileContent(i));
                If fileContent(i) = 10 Then Debug.Print ""
            Next i
        End If
    Else
        Debug.Print testFile & " is not in archive."
    End If
    Set zip = Nothing
End Sub


' simple function lets user select and order pdf files
Sub TestOrderPDFs()
    Dim files() As String
    files = PickFiles()
    Dim ufFileOrder As ufFileList: Set ufFileOrder = New ufFileList
    ufFileOrder.list = files
    ufFileOrder.Show
    files = ufFileOrder.list
    
    Dim i As Long
    For i = LBound(files) To UBound(files)
        Debug.Print files(i)
    Next i
End Sub

