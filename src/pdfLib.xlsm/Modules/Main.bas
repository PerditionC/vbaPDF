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
    Dim resultFn As String
    resultFn = SelectSaveFileName("C:\Users\jeremyd\Downloads\combined.pdf")
    If Not IsBlank(resultFn) Then
        CombinePDFs files, resultFn
    End If
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
    ' so we use combinedPdfDoc's nextObjId which automatically increments on each use, so we store for stable value during iteration
    Dim baseId As Long
    
    Dim ndx As Long
    For ndx = LBound(sourceFiles) To UBound(sourceFiles)
        Dim pdfDoc As pdfDocument: Set pdfDoc = New pdfDocument
        
        Debug.Print "Loading file " & ndx + 1 & " - " & sourceFiles(ndx)
        Application.StatusBar = ndx + 1 & " - " & sourceFiles(ndx)
        'loadPdf sourceFiles(ndx), trailer, xrefTable, info, root, pdfObjs
        pdfDoc.openPdf sourceFiles(ndx)
        
        ' adjust obj id's so no conflict with previously stored ones
        baseId = combinedPdfDoc.nextObjId
        pdfDoc.renumberIds baseId
            
        ' determine highest id used, 1st obj we add or next file will start at this + 1
        ' Note: we need to use pdfDoc.xrefTable's size and not combinedPdfDoc.xrefTable as we are reserving full count from just loaded pdf document
        baseId = baseId + pdfDoc.xrefTable.count - 1 ' highest id possible so far
        combinedPdfDoc.nextObjId = baseId + 1
            
        ' add the first page of this document as a Named Destination to our combined pdf
        ' Note: we assume current pageCount is how many existing pages there are, with
        ' the next page being 1st of just loaded document, then subtract 1 as page# begins with 0
        Dim docDestinationName As pdfValue
        Set docDestinationName = pdfValue.NewNameValue("/" & pdfDoc.Title & ".1", utf8BOM:=False)
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
        
        ' we add a bookmark to 1st page of each combined document
        ' TODO, make this optional
        ' TODO, check if we already did this, i.e. combining a previously combined file and another
        ' defaults should include /First /Last along with /Count, /Title /Prev /Next and optionally /Dest
        Dim defaults As Dictionary: Set defaults = New Dictionary
        Set defaults("/Title") = pdfValue.NewValue(pdfDoc.Title)
        Set defaults("/Dest") = docDestinationName
        Dim parentOutlineItem As pdfValue
        Set parentOutlineItem = combinedPdfDoc.NewOutlineItem(combinedPdfDoc.Outlines, defaults)
        combinedPdfDoc.AddOutlineItem Nothing, parentOutlineItem
        ' don't save here as we may need to adjust /Prev & /Next values
        
        ' next we need to merge any existing bookmarks
        If pdfDoc.rootCatalog.hasKey("/Outlines") Then
            ' Note: these are objs in pdfDoc so we need to remove if we add equivalent ones to combindedPdfDoc to avoid duplicates
            Dim outline As pdfValue
            Dim topLevelOutline As pdfValue
            Set topLevelOutline = pdfDoc.rootCatalog.asDictionary().item("/Outlines")
            If topLevelOutline.valueType <> PDF_ValueType.PDF_Null Then
                If topLevelOutline.valueType = PDF_ValueType.PDF_Reference Then
                    Set topLevelOutline = pdfDoc.getObject(topLevelOutline.value, topLevelOutline.generation)
                End If
                pdfDoc.objectCache.Remove topLevelOutline.ID
                
                ' toplevel outline probably doesn't have a title or siblings, so adds useless indirection
                If topLevelOutline.asDictionary().Exists("/Next") Or topLevelOutline.asDictionary().Exists("/Title") Then
                    ' we need to change its id so doesn't conflict
                    topLevelOutline.ID = -1
                    combinedPdfDoc.AddOutlineItem parentOutlineItem, topLevelOutline
                    ' we cycle through topLevelOutline, but use parentOutlineItem as its direct descendant's parent object, so skipping empty level
                    Set parentOutlineItem = topLevelOutline
                'Else skip adding it, and adjust its kids to point to out new filelevel bookmark as their parent
                End If
                
                ' TODO, if toplevel does has siblings, we need to handle them somehow, for now we just ignore and drop them
                                    
                ' now we need to update all its kids (but not add)
                If topLevelOutline.asDictionary().Exists("/First") Then
                    Set outline = topLevelOutline.asDictionary().item("/First")
                    Dim nextOutline As pdfValue
                    Dim firstOutline As pdfValue
                    Dim prevOutline As pdfValue
                    Do While Not outline Is Nothing
                        If outline.valueType = PDF_ValueType.PDF_Reference Then Set outline = pdfDoc.getObject(outline.value, outline.generation)
                        pdfDoc.objectCache.Remove outline.ID
                        outline.ID = combinedPdfDoc.nextObjId
                        outline.generation = 0
                        Set outline.asDictionary("/Parent") = parentOutlineItem.referenceObj
                        combinedPdfDoc.objectCache.Add outline.ID, outline
                        If firstOutline Is Nothing Then
                            Set firstOutline = outline ' really just a flag, could use a Boolean here
                            ' update parent's first
                            Set parentOutlineItem.asDictionary("/First") = outline.referenceObj
                        End If
                        If Not prevOutline Is Nothing Then
                            ' update next and prev links
                            Set prevOutline.asDictionary("/Next") = outline.referenceObj
                            Set outline.asDictionary("/Prev") = prevOutline.referenceObj
                        End If
                        ' TODO recurse from here too!
                        If outline.asDictionary().Exists("/Next") Then
                            Set prevOutline = outline
                            Set outline = outline.asDictionary().item("/Next")
                        Else
                            ' update parent's last
                            Set parentOutlineItem.asDictionary("/Last") = outline.referenceObj
                            Set outline = Nothing
                        End If
                    Loop
                End If
            End If
        Else
        End If
            
        ' remove /Root object and ensure only left with 1 /Root
        ' Warning: objectCache is used to convert references to objs, so do not attempt to retrieve any
        ' objects via obj reference after removing them from objectCache
        pdfDoc.objectCache.Remove pdfDoc.rootCatalog.ID
        ' also need to remove /Info from cache
        If pdfDoc.objectCache.Exists(pdfDoc.Info.ID) Then
            pdfDoc.objectCache.Remove pdfDoc.Info.ID
        End If
        ' and now save all the pages and other non-toplevel objects to our combined document
        combinedPdfDoc.SavePdfObjects outputFileNum, pdfDoc.objectCache, offset
        
        DoEvents
    Next ndx
    
    ' save outline objects
    combinedPdfDoc.SaveOutlinePdfObjects outputFileNum, combinedPdfDoc.Outlines, offset
    
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


'Create a FileDialog object as a File Picker dialog box and returns String array of files selected.
Function SelectSaveFileName(ByVal suggestedPath As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    With fd
        .AllowMultiSelect = False
        .Title = "Save PDF file as:"
        .InitialFileName = suggestedPath
        ' .Filters can not be modified
        '.FilterIndex = 26  ' PDF
        Dim ndx As Long
        For ndx = 1 To .Filters.count
            If InStr(1, .Filters(ndx).Extensions, "*.pdf", vbTextCompare) > 0 Then
                .FilterIndex = ndx
                Exit For
            End If
        Next
        Debug.Print .FilterIndex & " " & .Filters(ndx).Description & "-" & .Filters(ndx).Extensions
 
        'Use the Show method to display the File Picker dialog box and return the user's action.
        'The user pressed the button.
        If .Show = -1 Then
            If .SelectedItems.count > 0 Then
                'Step through each string in the FileDialogSelectedItems collection.
                Dim files() As String
                ReDim files(0 To .SelectedItems.count - 1)
                'Dim ndx As Long
                For ndx = 1 To .SelectedItems.count
                    files(ndx - 1) = .SelectedItems(ndx)
                Next ndx
                SelectSaveFileName = files(0)
            Else
                'SelectSaveFileName = vbNullString
            End If
        Else 'The user pressed Cancel.
            'SelectSaveFileName = vbNullString
        End If
        
        '.Filters.Clear
    End With
 
    Set fd = Nothing
End Function

