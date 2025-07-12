VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufBookmarkEditor 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ufBookmarkEditor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufBookmarkEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PDF Bookmark Editor Form
Option Explicit


Private WithEvents lstBookmarks As MSForms.ListBox
Attribute lstBookmarks.VB_VarHelpID = -1
Private WithEvents txtName As MSForms.TextBox
Attribute txtName.VB_VarHelpID = -1
Private WithEvents txtTitle As MSForms.TextBox
Attribute txtTitle.VB_VarHelpID = -1
Private WithEvents txtDest As MSForms.TextBox
Attribute txtDest.VB_VarHelpID = -1
Private WithEvents cmdUpdate As MSForms.CommandButton
Attribute cmdUpdate.VB_VarHelpID = -1
Private WithEvents cmdSave As MSForms.CommandButton
Attribute cmdSave.VB_VarHelpID = -1
Private lblName As MSForms.Label
Private lblTitle As MSForms.Label
Private lblDest As MSForms.Label

Private pdfDoc As pdfDocument
Private currentBookmarkIndex As Long

Public Function CreateBookmarkEditorForm(ByRef pdfDoc As pdfDocument) As ufBookmarkEditor
    ' Create the UserForm at runtime
    Set CreateBookmarkEditorForm = New ufBookmarkEditor
    
    With CreateBookmarkEditorForm
        .Caption = "PDF Bookmark Editor"
        .Width = 640 ' 640 pixels
        .Height = 480 ' 480 pixels
        .StartUpPosition = 1 ' CenterOwner
        
        ' Create controls
        .CreateControls
    
        ' Initialize form
        .InitializeForm
        
        ' update with document info
        .SetPDFDocument pdfDoc
    
        ' Show the form
        .Show
    End With
End Function

Friend Sub CreateControls()
    ' Create ListBox for bookmarks
    Set lstBookmarks = Me.Controls.Add("Forms.ListBox.1", "lstBookmarks")
    With lstBookmarks
        .Left = 18 ' ~240 twips converted to points
        .Top = 18
        .Width = 585 ' ~9135 twips converted to points
        .Height = 315 ' ~4935 twips converted to points
        .ColumnCount = 3
        .ColumnWidths = "100 pt;200 pt;200 pt"
        .ColumnHeads = True
    End With
    
    ' Create Name Label
    Set lblName = Me.Controls.Add("Forms.Label.1", "lblName")
    With lblName
        .Caption = "Name:"
        .Left = 18
        .Top = 345
        .Width = 75
        .Height = 15
    End With
    
    ' Create Name TextBox
    Set txtName = Me.Controls.Add("Forms.TextBox.1", "txtName")
    With txtName
        .Left = 100
        .Top = 342
        .Width = 520
        .Height = 18
    End With
    
    ' Create Title Label
    Set lblTitle = Me.Controls.Add("Forms.Label.1", "lblTitle")
    With lblTitle
        .Caption = "Title:"
        .Left = 18
        .Top = 370
        .Width = 75
        .Height = 15
    End With
    
    ' Create Title TextBox
    Set txtTitle = Me.Controls.Add("Forms.TextBox.1", "txtTitle")
    With txtTitle
        .Left = 100
        .Top = 367
        .Width = 520
        .Height = 18
    End With
    
    ' Create Dest Label
    Set lblDest = Me.Controls.Add("Forms.Label.1", "lblDest")
    With lblDest
        .Caption = "Dest:"
        .Left = 18
        .Top = 395
        .Width = 75
        .Height = 15
    End With
    
    ' Create Dest TextBox
    Set txtDest = Me.Controls.Add("Forms.TextBox.1", "txtDest")
    With txtDest
        .Left = 100
        .Top = 392
        .Width = 520
        .Height = 18
    End With
    
    ' Create Update Button
    Set cmdUpdate = Me.Controls.Add("Forms.CommandButton.1", "cmdUpdate")
    With cmdUpdate
        .Caption = "Update"
        .Left = 450
        .Top = 420
        .Width = 80
        .Height = 25
    End With
    
    ' Create Save Button
    Set cmdSave = Me.Controls.Add("Forms.CommandButton.1", "cmdSave")
    With cmdSave
        .Caption = "Save PDF"
        .Left = 540
        .Top = 420
        .Width = 80
        .Height = 25
    End With
End Sub

Friend Sub InitializeForm()
    ' Initialize form
    currentBookmarkIndex = -1
    
    ' Clear text fields
    ClearFields
    
    ' Load PDF if available
    LoadPDFBookmarks
End Sub

Private Sub LoadPDFBookmarks()
    On Error GoTo ErrorHandler
    
    ' This would typically be called after loading a PDF document
    ' For now, we'll assume pdfDoc is already set
    If Not pdfDoc Is Nothing Then
        PopulateBookmarksList
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error loading bookmarks: " & Err.Description, vbCritical
    Resume
End Sub

'================================================================================
'  BOOKMARK-TREE  ?  LISTBOX
'================================================================================
Private Sub PopulateBookmarksList()
    On Error GoTo ErrHandler
    
    lstBookmarks.Clear
    
    If pdfDoc Is Nothing Then Exit Sub
    
    Dim outlineDict As Dictionary
    Set outlineDict = pdfDoc.Outlines.asDictionary()
    If outlineDict Is Nothing Then Exit Sub
    
    'Top level always starts at /First
    If outlineDict.Exists("/First") Then
        Dim firstRef As pdfValue
        Set firstRef = outlineDict("/First")          '<<  12 0 R  (reference)
        
        Dim firstObj As pdfValue
        Set firstObj = pdfDoc.getObject(firstRef.value, firstRef.generation)  '<< actual object
        AddBookmarkRecursive firstObj, 0                 'level = 0
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "PopulateBookmarksList ? " & Err.Description, vbExclamation
End Sub


'-------------------------------------------------------------------------------
' Recursively walks the bookmark linked list:
'   nodeDict(/First)  –> first child
'   nodeDict(/Next)   –> next sibling
'-------------------------------------------------------------------------------
Private Sub AddBookmarkRecursive(ByVal bmObj As pdfValue, ByVal level As Long)
    On Error GoTo ErrHandler:

    If bmObj Is Nothing Then Exit Sub
    
    Dim nodeDict As Dictionary
    Set nodeDict = bmObj.asDictionary()
    If nodeDict Is Nothing Then Exit Sub
    
    '----- 1. Read the fields we care about ------------------------------------
    Dim rowName As String, rowTitle As String, rowDest As String
    
    rowName = bmObj.ID & " " & bmObj.generation & " R"                             'e.g.  "12 0 R"
    
    If nodeDict.Exists("/Title") Then _
        rowTitle = CStr(nodeDict("/Title").value)
    If nodeDict.Exists("/Dest") Then _
        rowDest = CStr(nodeDict("/Dest").value)
    
    'Indent title to show hierarchy
    rowTitle = String(level * 2, " ") & rowTitle
    
    '----- 2. Push the row into the ListBox ------------------------------------
    With lstBookmarks
        .AddItem rowName                      'col 0
        Dim r&: r = .ListCount - 1
        .list(r, 1) = rowTitle               'col 1
        .list(r, 2) = rowDest                'col 2
    End With
    
    '----- 3. Recurse into children --------------------------------------------
    If nodeDict.Exists("/First") Then
        Dim childRef As pdfValue, childObj As pdfValue
        Set childRef = nodeDict("/First")
        Set childObj = pdfDoc.getObject(childRef.value, childRef.generation)
        AddBookmarkRecursive childObj, level + 1
    End If
    
    '----- 4. Walk to next sibling ---------------------------------------------
    If nodeDict.Exists("/Next") Then
        Dim nextRef As pdfValue, nextObj As pdfValue
        Set nextRef = nodeDict("/Next")
        Set nextObj = pdfDoc.getObject(nextRef.value, nextRef.generation)
        AddBookmarkRecursive nextObj, level          'same depth
    End If
    
    Exit Sub
ErrHandler:
    Debug.Print Err.Description & " (" & Err.Number & ")"
    Stop
    Resume
End Sub

Private Function GetDictValue(dict As Dictionary, keyName As String) As String
    If dict.Exists(keyName) Then
        GetDictValue = CStr(dict(keyName))
    Else
        GetDictValue = ""
    End If
End Function

Private Sub lstBookmarks_Click()
    If lstBookmarks.ListIndex >= 0 Then
        currentBookmarkIndex = lstBookmarks.ListIndex
        LoadSelectedBookmark
    End If
End Sub

Private Sub LoadSelectedBookmark()
    On Error GoTo ErrHandler
    
    If currentBookmarkIndex < 0 Or lstBookmarks.ListCount = 0 Then Exit Sub
    
    ' Get the bookmark reference from the selected row (column 0 contains the object reference)
    Dim bookmarkRef As String
    bookmarkRef = lstBookmarks.list(currentBookmarkIndex, 0)  ' e.g. "12 0 R"
    
    If Len(bookmarkRef) = 0 Then Exit Sub
    
    ' Get the actual bookmark object using the reference
    Dim bookmarkObj As pdfValue
    Dim objId As Long, objGen As Long
    objId = Left(bookmarkRef, InStr(1, bookmarkRef, " ", vbBinaryCompare))
    objGen = Mid(bookmarkRef, 1 + InStr(1, bookmarkRef, " ", vbBinaryCompare), 1) ' *** FIXME
    Set bookmarkObj = pdfDoc.getObject(objId, objGen)
    
    If bookmarkObj Is Nothing Then Exit Sub
    
    ' Get the dictionary for this bookmark
    Dim bookmarkDict As Dictionary
    Set bookmarkDict = bookmarkObj.asDictionary()
    
    If bookmarkDict Is Nothing Then Exit Sub
    
    ' Populate text fields with bookmark data
    txtName.Text = bookmarkRef  ' The object reference itself
    
    ' Extract title (remove indentation spaces added during display)
    Dim rawTitle As String
    If bookmarkDict.Exists("/Title") Then
        rawTitle = CStr(bookmarkDict("/Title").value)
        txtTitle.Text = Trim(rawTitle)  ' Remove any leading/trailing spaces
    Else
        txtTitle.Text = vbNullString
    End If
    
    ' Extract destination
    If bookmarkDict.Exists("/Dest") Then
        txtDest.Text = CStr(bookmarkDict("/Dest").value)
    Else
        txtDest.Text = vbNullString
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "LoadSelectedBookmark ? " & Err.Description, vbExclamation
    ClearFields
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo ErrorHandler
    
    If currentBookmarkIndex < 0 Or lstBookmarks.ListCount = 0 Then
        MsgBox "Please select a bookmark to update.", vbInformation
        Exit Sub
    End If
    
    ' Get the bookmark reference from the selected row (column 0 contains the object reference)
    Dim bookmarkRef As String
    bookmarkRef = lstBookmarks.list(currentBookmarkIndex, 0)  ' e.g. "12 0 R"
    
    If Len(bookmarkRef) = 0 Then
        MsgBox "Invalid bookmark reference.", vbExclamation
        Exit Sub
    End If
    
    ' Get the actual bookmark object using the reference
    Dim bookmarkObj As pdfValue
    Dim objId As Long, objGen As Long
    objId = Left(bookmarkRef, InStr(1, bookmarkRef, " ", vbBinaryCompare))
    objGen = Mid(bookmarkRef, 1 + InStr(1, bookmarkRef, " ", vbBinaryCompare), 1) ' *** FIXME
    Set bookmarkObj = pdfDoc.getObject(objId, objGen)
    
    If bookmarkObj Is Nothing Then
        MsgBox "Could not retrieve bookmark object.", vbExclamation
        Exit Sub
    End If
    
    ' Get the dictionary for this bookmark
    Dim bookmarkDict As Dictionary
    Set bookmarkDict = bookmarkObj.asDictionary()
    
    If bookmarkDict Is Nothing Then
        MsgBox "Could not access bookmark dictionary.", vbExclamation
        Exit Sub
    End If
    
    ' Update dictionary values with text field contents
    ' Note: We don't update the object reference itself (Name field)
    
    If Len(Trim(txtTitle.Text)) > 0 Then
        ' Update the /Title key in the dictionary
        bookmarkDict("/Title").value = Trim(txtTitle.Text)
    End If
    
    If Len(Trim(txtDest.Text)) > 0 Then
        ' Update the /Dest key in the dictionary
        ' Note: Destination format may vary (array, string, etc.)
        bookmarkDict("/Dest").value = Trim(txtDest.Text)
    End If
    
    ' Refresh the entire bookmark list to show changes
    PopulateBookmarksList
    
    ' Restore selection if possible
    If currentBookmarkIndex < lstBookmarks.ListCount Then
        lstBookmarks.ListIndex = currentBookmarkIndex
    End If
    
    MsgBox "Bookmark updated successfully!", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error updating bookmark: " & Err.Description, vbCritical
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler
    
    If pdfDoc Is Nothing Then
        MsgBox "No PDF document loaded.", vbInformation
        Exit Sub
    End If
    
    Dim fileName As String
    fileName = Application.GetSaveAsFilename( _
        InitialFileName:="BookmarksUpdated.pdf", _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Save PDF Document")
    
    If fileName <> "False" Then
        ' Save the PDF document
        pdfDoc.savePdfAs fileName
        MsgBox "PDF saved successfully to: " & fileName, vbInformation
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error saving PDF: " & Err.Description, vbCritical
End Sub

Private Sub ClearFields()
    If Not txtName Is Nothing Then txtName.Text = vbNullString
    If Not txtTitle Is Nothing Then txtTitle.Text = vbNullString
    If Not txtDest Is Nothing Then txtDest.Text = vbNullString
End Sub

' Public method to set the PDF document object
Public Sub SetPDFDocument(pdfObject As Object)
    Set pdfDoc = pdfObject
    LoadPDFBookmarks
End Sub

Public Sub CleanUp()
    ' Clean up object references
    Set lstBookmarks = Nothing
    Set txtName = Nothing
    Set txtTitle = Nothing
    Set txtDest = Nothing
    Set cmdUpdate = Nothing
    Set cmdSave = Nothing
    Set lblName = Nothing
    Set lblTitle = Nothing
    Set lblDest = Nothing
    Set pdfDoc = Nothing
    Set ufBookmarkEditor = Nothing
End Sub

Private Sub UserForm_Terminate()
    CleanUp
End Sub
