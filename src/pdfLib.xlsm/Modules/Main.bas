Attribute VB_Name = "Main"
Option Explicit


Public Sub CombinePDFs(ByRef sourceFiles() As String, ByRef outFile As String)
    'Dim oPages As pdfValue  ' /Type /Pages with /Count # /Kids [ /Page references ...]
    Dim combinedPdfDoc As pdfDocument: Set combinedPdfDoc = New pdfDocument
    ' initialize with some basic structures
    combinedPdfDoc.AddInfo
    combinedPdfDoc.AddPages
    combinedPdfDoc.AddOutlines
    With combinedPdfDoc.rootCatalog.asDictionary()
        Set .Item("/PageMode") = pdfValue.NewNameValue("/UseOutlines") ' default is "/UseNone"
    End With
    
    Dim offset As Long
    Dim outputFileNum As Integer
    outputFileNum = combinedPdfDoc.SavePdfHeader(outFile, offset)
    
    Dim baseId As Long
    ' we don't yet know the id's used by our /Root, /Info and top level /Pages objs
    With combinedPdfDoc
        .rootCatalog.id = -1
        .Info.id = -2
        .pages.id = -3
    
        baseId = .nextObjId
    End With
    
    Dim ndx As Long
    For ndx = LBound(sourceFiles) To UBound(sourceFiles)
        Dim pdfDoc As pdfDocument: Set pdfDoc = New pdfDocument
        
        Debug.Print "Loading file " & ndx + 1 & " - " & sourceFiles(ndx)
        Application.StatusBar = ndx + 1 & " - " & sourceFiles(ndx)
        'loadPdf sourceFiles(ndx), trailer, xrefTable, info, root, pdfObjs
        pdfDoc.loadPdf sourceFiles(ndx)
        pdfDoc.parsePdf
        
        ' for each additional document we need to update /Pages
        Dim pages As pdfValue
        Set pages = pdfDoc.pages()
        
        ' since we are about to remove it, we use 1st pdf doc's /Root id for our new top level /Pages (so we have for /Parent references)
        ' Note: once we add pdfDoc.pages to combinedPdfDoc.pages and save, we can no longer change id of combinedPdfDoc.pages
        With combinedPdfDoc
            If .pages.id < 0 Then
                .pages.id = pdfDoc.rootCatalog.id
            End If
        End With
        
        ' remove /Root object, we need to copy/merge some optional fields such as /Outline for bookmarks
        ' and ensure only left with 1 /Root
        pdfDoc.objectCache.Remove pdfDoc.rootCatalog.id
        ' also need to remove /Info from cache
        If pdfDoc.objectCache.Exists(pdfDoc.Info.id) Then
            pdfDoc.objectCache.Remove pdfDoc.Info.id
        End If
        
        ' we inject a new top level /Pages object which we add all the document /Pages to
        ' so we need to add this /Pages to our top level /Pages and add a /Parent indirect reference
        combinedPdfDoc.AddPages pdfDoc.pages
        
        If pdfDoc.rootCatalog.hasKey("/Outlines") Then
            ' hack for now, just copy over
            combinedPdfDoc.rootCatalog.asDictionary().Add "/Outlines", pdfDoc.rootCatalog.asDictionary().Item("/Outlines")
        End If
            
        combinedPdfDoc.SavePdfObjects outputFileNum, pdfDoc.objectCache, offset
        
        ' determine highest id used, 1st obj in next file will start at this + 1
        ' Note: we need to use pdfDoc.xrefTable's size and not combinedPdfDoc.xrefTable as we are reserving full count from just loaded pdf document
        baseId = baseId + pdfDoc.xrefTable.Count - 1 ' highest id possible so far
        combinedPdfDoc.nextObjId = baseId + 1
        DoEvents
    Next ndx
    
    ' we need to set valid id's for our top level objs
    With combinedPdfDoc
        .Info.id = .nextObjId
        .rootCatalog.id = .nextObjId
    End With
    
    ' save updated /Pages object (but not nested objects as already saved)
    combinedPdfDoc.SavePdfObject outputFileNum, combinedPdfDoc.pages, offset
    combinedPdfDoc.SavePdfObject outputFileNum, combinedPdfDoc.rootCatalog, offset
    
    ' writes out trailer and cross reference table
    combinedPdfDoc.SavePdfTrailer outputFileNum, offset
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


'Create a FileDialog object as a File Picker dialog box and returns String array of files selected.
Function PickFiles(Optional ByVal AllowMultiSelect As Boolean = True) As String()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = AllowMultiSelect
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
