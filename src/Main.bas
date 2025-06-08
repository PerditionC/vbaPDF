Attribute VB_Name = "Main"
Option Explicit


#If False Then
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
    Dim C As Collection
    
    Set d = pages1.Value.Value
    Set obj = d.Item("/Kids")
    Set C = obj.Value
    For Each v In C
        Set obj = v ' obj reference
        kids.Add obj
    Next v
    
    ' Warning! TODO! need to fixup actual obj Referenced in here so its /Parent refers back to our pages.id obj
    Set d = pages2.Value.Value
    Set obj = d.Item("/Kids")
    Set C = obj.Value
    For Each v In C
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
#End If


Public Sub CombinePDFs(ByRef sourceFiles() As String, ByRef outFile As String)
    'Dim oPages As pdfValue  ' /Type /Pages with /Count # /Kids [ /Page references ...]
    Dim combinedPdfDoc As pdfDocument: Set combinedPdfDoc = New pdfDocument
    
    Dim offset As Long
    Dim outputFileNum As Integer
    outputFileNum = combinedPdfDoc.SavePdfHeader(outFile, offset)
    
    Dim ndx As Long, baseId As Long
    For ndx = LBound(sourceFiles) To UBound(sourceFiles)
        Dim pdfDoc As pdfDocument: Set pdfDoc = New pdfDocument
        
        Debug.Print "Loading file " & ndx + 1 & " - " & sourceFiles(ndx)
        Application.StatusBar = ndx + 1 & " - " & sourceFiles(ndx)
        'loadPdf sourceFiles(ndx), trailer, xrefTable, info, root, pdfObjs
        pdfDoc.loadPdf sourceFiles(ndx)
        pdfDoc.parsePdf
        
        ' for each additional document we need to update /Pages
        'Set pages = FindPages(root, pdfObjs)
        Dim pages As pdfValue
        Set pages = pdfDoc.pages()
        
        If combinedPdfDoc.rootCatalog.valueType = PDF_ValueType.PDF_Null Then
            ' first time through we can just use /Root from pdf unchanged
            Set combinedPdfDoc.rootCatalog = pdfDoc.rootCatalog
            Set combinedPdfDoc.trailer = pdfDoc.trailer
            Set combinedPdfDoc.Info = pdfDoc.Info
            Set combinedPdfDoc.pages = pdfDoc.pages
            
            ' so remove pages object for now, we add a single one back after we have gone through all of them
            pdfDoc.objectCache.Remove pages.id
        Else
            ' each additional pdf need to remove /Root, so only left with 1 (the first) /Root
            pdfDoc.objectCache.Remove pdfDoc.rootCatalog.id
#If False Then
            ' so remove pages object for now, we add a single one back after we have gone through all of them
            pdfDoc.objectCache.Remove pages.id
            Set combinedPdfDoc.pages = CombinePages(oPages, pages, baseId)
#Else
            ' instead of removing /Pages and combining into a single /Pages, we add a higher level /Pages
            ' that points to all /Pages, but we need to fixup their parents to point to our new /Pages
            ' so we use 1st /Pages as new parent for all future /Pages
            Dim obj As pdfValue
            Set obj = New pdfValue
            obj.valueType = PDF_ValueType.PDF_Reference
            obj.Value = combinedPdfDoc.pages.id - baseId  ' Warning! this may be negative, but corrects when Saved and baseId added back!
            'If pages.Value.Value.Exists("/Parent") Then
            If pages.asDictionary.Exists("/Parent") Then
                ' we don't handle this case, it shouldn't exist
                Stop
            Else
                Set pages.asDictionary.Item("/Parent") = obj ' should be Reference obj
            End If
            ' and add to our top level /Pages
            Set obj = combinedPdfDoc.pages.asDictionary.Item("/Kids")
            Dim C As Collection
            Set C = obj.Value
            ' but as a reference
            Set obj = New pdfValue
            obj.valueType = PDF_ValueType.PDF_Reference
            obj.Value = pages.id + baseId   ' we must correct id here as we don't update when saving this obj
            C.Add obj
            ' we also need to update our count
            Set obj = combinedPdfDoc.pages.asDictionary.Item("/Count")
            ' ### not +1 as not a count of /Kids but /Count of total pages, so need to sum child counts
            'obj.value = CLng(obj.value + 1)
            Dim Count As Long
            Count = CLng(pages.asDictionary.Item("/Count").Value)
            obj.Value = CLng(obj.Value + Count)
#End If
        End If
        
        combinedPdfDoc.SavePdfObjects outputFileNum, pdfDoc.objectCache, offset, baseId
        
        ' determine highest id used, 1st obj in next file will start at this + 1
        ' Note: we need to use pdfDoc.xrefTable's size and not combinedPdfDoc.xrefTable as we are reserving full count from just loaded pdf document
        baseId = baseId + pdfDoc.xrefTable.Count - 1 ' highest id possible so far
        DoEvents
    Next ndx
    
    ' save updated /Pages object (but not nested objects as already saved)
    combinedPdfDoc.SavePdfObject outputFileNum, combinedPdfDoc.pages, offset
    
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
Function PickFiles() As String()
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
