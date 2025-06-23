Attribute VB_Name = "Main"
Option Explicit


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


Public Sub CombinePDFs(ByRef sourceFiles() As String, ByRef outFile As String)
    Dim combinedPdfDoc As pdfDocument: Set combinedPdfDoc = New pdfDocument
    ' initialize with some basic structures
    With combinedPdfDoc
        .AddInfo
        .AddPages
        .AddOutlines
        With .rootCatalog.asDictionary()
            Set .item("/PageMode") = pdfValue.NewNameValue("/UseOutlines") ' default is "/UseNone"
        End With
    End With
    
    Dim offset As Long
    Dim outputFileNum As Integer
    outputFileNum = combinedPdfDoc.SavePdfHeader(outFile, offset)
    
    ' get where to start renumbering objs in pdf, we need to skip past our toplevel /Root, /Info, /Pages, & /Outlines
    Dim baseId As Long: baseId = combinedPdfDoc.nextObjId
    
    Dim ndx As Long
    For ndx = LBound(sourceFiles) To UBound(sourceFiles)
        Dim pdfDoc As pdfDocument: Set pdfDoc = New pdfDocument
        
        Debug.Print "Loading file " & ndx + 1 & " - " & sourceFiles(ndx)
        Application.StatusBar = ndx + 1 & " - " & sourceFiles(ndx)
        'loadPdf sourceFiles(ndx), trailer, xrefTable, info, root, pdfObjs
        pdfDoc.loadPdf sourceFiles(ndx)
        pdfDoc.parsePdf
        
        ' adjust obj id's so no conflict with previously stored ones
        pdfDoc.renumberIds baseId + 1   ' +1 as baseId used for our /OutlineItem
            
        ' add the first page of this document as a Named Destination to our combined pdf
        ' Note: we assume current pageCount is how many existing pages there are, with
        ' the next page being 1st of just loaded document, then subtract 1 as page# begins with 0
        Dim docDestinationName As pdfValue
        Set docDestinationName = pdfValue.NewNameValue("/" & pdfDoc.Title & ".1", utf8:=True)
        Dim docDestination As pdfValue
        Set docDestination = combinedPdfDoc.NewDestination(combinedPdfDoc.pageCount, PDF_FIT.PDF_FIT)
        combinedPdfDoc.AddNamedDestinations docDestinationName, docDestination
            
        ' copy over any pre-existing Named Destinations
        ' Note: these are in a <<dictionary>> in /Root/Pages so not automatically included
        If pdfDoc.Dests.valueType = PDF_ValueType.PDF_Dictionary Then
            Dim dict As Dictionary: Set dict = pdfDoc.Dests.asDictionary()
            If Not dict Is Nothing Then
                Dim v As Variant
                For Each v In dict.Keys
                    ' we treat name as PDF_String instead of PDF_Name to avoid potential double escaping
                    ' and because while PDF_Name recommended, not required so could just be a PDF_String
                    Dim name As pdfValue: Set name = pdfValue.NewValue(v)
                    combinedPdfDoc.AddNamedDestinations name, dict(v)
                Next v
                Set v = Nothing
            End If
            Set dict = Nothing
        End If
        
        ' we inject a new top level /Pages object which we add all the document /Pages to
        ' so we need to add this /Pages to our top level /Pages and add a /Parent indirect reference
        combinedPdfDoc.AddPages pdfDoc.Pages
        
        ' we need to copy/merge some optional fields such as /Outline for bookmarks
        If False And pdfDoc.rootCatalog.hasKey("/Outlines") Then
            ' Note: these are objs in pdfDoc so we need to remove if we add equivalent ones to combindedPdfDoc to avoid duplicates
            ' hack for now, just copy over
            'combinedPdfDoc.rootCatalog.asDictionary().Add "/Outlines", pdfDoc.rootCatalog.asDictionary().item("/Outlines")
            ' TODO
        Else
            ' defaults should include /First /Last along with /Count, /Title /Prev /Next and optionally /Dest
            Dim defaults As Dictionary: Set defaults = New Dictionary
            Set defaults("/Title") = pdfValue.NewValue(pdfDoc.Title)
            Set defaults("/Dest") = docDestinationName
            Dim outlineItem As pdfValue
            Set outlineItem = combinedPdfDoc.NewOutlineItem(combinedPdfDoc.Outlines, defaults)
            combinedPdfDoc.AddOutlines outlineItem
            ' don't save here as we may need to adjust /Prev & /Next values
        End If
            
        ' remove /Root object and ensure only left with 1 /Root
        ' Warning: objectCache is used to convert references to objs, so do not attempt to retrieve any
        ' objects via obj reference after removing them from objectCache
        pdfDoc.objectCache.Remove pdfDoc.rootCatalog.id
        ' also need to remove /Info from cache
        If pdfDoc.objectCache.Exists(pdfDoc.Info.id) Then
            pdfDoc.objectCache.Remove pdfDoc.Info.id
        End If
        ' and now save all the pages and other non-toplevel objects to our combined document
        combinedPdfDoc.SavePdfObjects outputFileNum, pdfDoc.objectCache, offset
        
        ' determine highest id used, 1st obj in next file will start at this + 1
        ' Note: we need to use pdfDoc.xrefTable's size and not combinedPdfDoc.xrefTable as we are reserving full count from just loaded pdf document
        baseId = baseId + pdfDoc.xrefTable.count - 1 ' highest id possible so far
        combinedPdfDoc.nextObjId = baseId + 1
        DoEvents
    Next ndx
    
    combinedPdfDoc.SavePdfObject outputFileNum, combinedPdfDoc.Outlines, offset
    If combinedPdfDoc.Outlines.hasKey("/First") Then
        Dim outlineObj As pdfValue
        Set outlineObj = combinedPdfDoc.Outlines.asDictionary().item("/First")
        Set outlineObj = combinedPdfDoc.getObject(outlineObj.value, outlineObj.generation)
        Do While Not outlineObj Is Nothing
            combinedPdfDoc.SavePdfObject outputFileNum, outlineObj, offset
            If outlineObj.hasKey("/Next") Then
                Set outlineObj = outlineObj.asDictionary().item("/Next")
                Set outlineObj = combinedPdfDoc.getObject(outlineObj.value, outlineObj.generation)
            Else
                Set outlineObj = Nothing
            End If
        Loop
    End If
    ' save updated /Pages object (but not nested objects as already saved)
    combinedPdfDoc.SavePdfObject outputFileNum, combinedPdfDoc.Pages, offset
    combinedPdfDoc.SavePdfObject outputFileNum, combinedPdfDoc.rootCatalog, offset
    
    ' writes out trailer and cross reference table
    combinedPdfDoc.SavePdfTrailer outputFileNum, offset
    'SavePdf outFile, trailer, oXrefTable, info, root, oPdfObjs
    Debug.Print "Saved " & outFile
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
            ReDim files(0 To .SelectedItems.count - 1)
            Dim ndx As Long
            For ndx = 1 To .SelectedItems.count
                files(ndx - 1) = .SelectedItems(ndx)
            Next ndx
            PickFiles = files
        Else 'The user pressed Cancel.
        End If
        
        .Filters.Clear
    End With
 
    Set fd = Nothing
End Function
