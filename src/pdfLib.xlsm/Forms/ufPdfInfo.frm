VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPdfInfo 
   Caption         =   "View PDF Information"
   ClientHeight    =   12672
   ClientLeft      =   156
   ClientTop       =   576
   ClientWidth     =   13500
   OleObjectBlob   =   "ufPdfInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufPdfInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pdfDoc As pdfDocument

Private Sub cbContinue_Click()
    Me.Hide
    Unload Me
End Sub


Private Sub UpdateInfo()
    lblPdfTitle.Caption = pdfDoc.Title  ' defaults to filename, not including path if not found
    
    lblAuthor.Caption = vbNullString
    lblCreationDate.Caption = vbNullString
    'lblModDate.Caption = vbNullString
    lblProducer.Caption = vbNullString
    lblSubject.Caption = vbNullString
    lbInfo.Clear
    
    
    If pdfDoc.Info.valueType <> PDF_ValueType.PDF_Null Then
        'If pdfDoc.Info.asDictionary().Exists("/Title") Then lblPdfTitle.Caption = pdfDoc.Info.asDictionary("/Title").Value
        
        If pdfDoc.Info.asDictionary().Exists("/Author") Then lblAuthor.Caption = pdfDoc.Info.asDictionary("/Author").Value
        If pdfDoc.Info.asDictionary().Exists("/CreationDate") Then lblCreationDate.Caption = pdfDoc.Info.asDictionary("/CreationDate").Value
        'If pdfDoc.Info.asDictionary().Exists("/ModDate") Then lblModDate.Caption = pdfDoc.Info.asDictionary("/ModDate").Value
        If pdfDoc.Info.asDictionary().Exists("/Producer") Then lblProducer.Caption = pdfDoc.Info.asDictionary("/Producer").Value
        If pdfDoc.Info.asDictionary().Exists("/Subject") Then lblSubject.Caption = pdfDoc.Info.asDictionary("/Subject").Value
        
        Dim dict As Dictionary
        Set dict = pdfDoc.Info.asDictionary
        Dim v As Variant
        Dim obj As pdfValue
        For Each v In dict.Keys
            lbInfo.AddItem CStr(v) & "=" & dict.Item(v).Value
        Next v
    End If
End Sub

Private Sub UpdatePages()
    lblPageCount = pdfDoc.pageCount
End Sub


Private Sub UserForm_Activate()
    ' auto prompt for file to load if not passed programatically beforehand
    If pdfDoc Is Nothing Then
        Dim files() As String
        files = PickFiles()
        ' slight hack, but Not files = -1 if user cancelled, but Ubound(files) will throw an error
        If (Not files) <> -1 Then Set pdfDoc = pdfDocument.pdfDocument(files(0))
    End If

    ' show info, but only if something actually loaded
    If Not pdfDoc Is Nothing Then
        lblFilename.Caption = pdfDoc.filepath & pdfDoc.filename
        UpdateInfo
        UpdatePages
    End If
End Sub
