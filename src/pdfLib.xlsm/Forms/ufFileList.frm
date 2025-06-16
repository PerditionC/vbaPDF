VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufFileList 
   Caption         =   "Combine PDF files"
   ClientHeight    =   10140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   OleObjectBlob   =   "ufFileList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public list As Variant  ' of full filenames and paths

Public pdfDocs As Dictionary

Private Sub swapOrder(ByVal ndx1 As Long, ByVal ndx2 As Long)
    If ndx1 < 0 Or ndx2 < 0 Or ndx1 >= lbFiles.ListCount Or ndx2 >= lbFiles.ListCount Then Exit Sub
    Dim tmpVal As String
    tmpVal = lbFiles.list(ndx1)
    lbFiles.list(ndx1) = lbFiles.list(ndx2)
    lbFiles.list(ndx2) = tmpVal
End Sub

Private Sub saveList()
    Dim fileList() As String
    ReDim fileList(0 To lbFiles.ListCount - 1)
    Dim i As Long
    For i = 0 To lbFiles.ListCount - 1
        fileList(i) = lbFiles.list(i)
    Next i
    list = fileList
End Sub

Private Sub cbContinue_Click()
    Me.Hide
    saveList
End Sub

Private Sub cbReverseOrder_Click()
    Dim i As Long
    For i = 0 To (lbFiles.ListCount / 2) - 1
        swapOrder i, lbFiles.ListCount - 1 - i
    Next i
End Sub

Private Sub ckbShowInfo_Click()
    If FrameInfo.Visible <> ckbShowInfo.value Then
        ' note this test is reversed from what we want because we are about to change it
        If Not FrameInfo.Visible Then
            ufFileList.Height = ufFileList.Height + FrameInfo.Height
        Else
            ufFileList.Height = ufFileList.Height - FrameInfo.Height
        End If
    End If
    
    FrameInfo.Visible = ckbShowInfo.value
End Sub

Private Sub lbFiles_Click()
    Dim filename As String
    filename = lbFiles.list(lbFiles.ListIndex)
    ' default caption of just filename selected
    lblPdfTitle.Caption = filename
    
    lblAuthor.Caption = vbNullString
    lblCreationDate.Caption = vbNullString
    'lblModDate.Caption = vbNullString
    lblProducer.Caption = vbNullString
    
    
    Dim pdfDoc As pdfDocument
    If pdfDocs.Exists(filename) Then
ShowPdfMeta:
        Set pdfDoc = pdfDocs(filename)
        ' default to filename, not including path
        'lblPdfTitle.Caption = pdfDoc.filename
        lblPdfTitle.Caption = pdfDoc.Title
        If pdfDoc.Info.valueType <> PDF_ValueType.PDF_Null Then
            'If pdfDoc.Info.hasKey("/Title") Then lblPdfTitle.Caption = pdfDoc.Info.asDictionary("/Title").Value
        
            If pdfDoc.Info.hasKey("/Author") Then lblAuthor.Caption = pdfDoc.Info.asDictionary("/Author").value
            If pdfDoc.Info.hasKey("/CreationDate") Then lblCreationDate.Caption = pdfDoc.Info.asDictionary("/CreationDate").value
            'If pdfDoc.Info.hasKey("/ModDate") Then lblModDate.Caption = pdfDoc.Info.asDictionary("/ModDate").Value
            If pdfDoc.Info.hasKey("/Producer") Then lblProducer.Caption = pdfDoc.Info.asDictionary("/Producer").value
        End If
    Else
        ' load pdf
        Set pdfDoc = New pdfDocument
        If pdfDoc.loadPdf(filename) Then
            pdfDocs.Add filename, pdfDoc
            GoTo ShowPdfMeta
        End If
        
        lblPdfTitle.Caption = lblPdfTitle.Caption & "*"
    End If
End Sub

Private Sub spUpDown_SpinDown()
    swapOrder lbFiles.ListIndex, lbFiles.ListIndex + 1
    If lbFiles.ListIndex < (lbFiles.ListCount - 1) Then lbFiles.ListIndex = lbFiles.ListIndex + 1
End Sub

Private Sub spUpDown_SpinUp()
    swapOrder lbFiles.ListIndex, lbFiles.ListIndex - 1
    If lbFiles.ListIndex > 0 Then lbFiles.ListIndex = lbFiles.ListIndex - 1
End Sub

Private Sub UserForm_Activate()
    FrameInfo.Visible = ckbShowInfo.value
    
    Dim i As Long
    For i = LBound(list) To UBound(list)
        lbFiles.AddItem list(i)
    Next
    lbFiles.ListIndex = 0
End Sub

Private Sub UserForm_Initialize()
    Dim testData() As String
    'testData = Split("file1,file2,file3,file4,file5", ",")
    testData = Split("C:\Users\jeremyd\Downloads\rewritten.pdf,C:\Users\jeremyd\Downloads\Combined.pdf,C:\Users\jeremyd\Downloads\Combined.good.pdf,C:\Users\jeremyd\Downloads\rewritten.good.pdf,C:\Users\jeremyd\Downloads\test3.pdf", ",")
    list = testData
    Set pdfDocs = New Dictionary
End Sub


