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
    Set pdfDoc = New pdfDocument
    Debug.Print pdfDoc.Version
    Set pdfDoc = New pdfDocument
    Debug.Print "[" & pdfDoc.Header & "]"
End Sub

Sub TestReWritePdf()
    Const basedir As String = "C:\Users\jeremyd\Downloads\"
#If 0 Then
    Dim trailer As pdfValue
    Dim xrefTableOriginal As Dictionary
    Dim Info As pdfValue
    Dim root As pdfValue
    Dim pdfObjs As Dictionary
    
    loadPdf basedir & "test2.pdf", trailer, xrefTableOriginal, Info, root, pdfObjs
    
    Dim xrefTable As Dictionary
    Set xrefTable = NewXrefTable()
    savePdf basedir & "rewritten.pdf", trailer, xrefTable, Info, root, pdfObjs
#Else
    Dim pdfDoc As pdfDocument
    Set pdfDoc = New pdfDocument
    If Not pdfDoc.loadPdf(basedir & "test2.pdf") Then
        Debug.Print "Error loading " & pdfDoc.filename
    End If
    If Not pdfDoc.parsePdf() Then
        Debug.Print "Error parsing pdf " & pdfDoc.filename
    End If
    If Not pdfDoc.savePdfAs(basedir & "rewritten.pdf") Then
        Debug.Print "Error saving " & "rewritten.pdf"
    End If
#End If
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
