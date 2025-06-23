VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPdfInfo 
   Caption         =   "View PDF Information"
   ClientHeight    =   6615
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   10800
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
    'pdfDoc.renumberIds pdfDoc.trailer, 100
    'pdfDoc.savePdfAs "renumberd.pdf"
    'Stop
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
        
        If pdfDoc.Info.hasKey("/Author") Then lblAuthor.Caption = pdfDoc.Info.asDictionary("/Author").value
        If pdfDoc.Info.hasKey("/CreationDate") Then lblCreationDate.Caption = pdfDoc.Info.asDictionary("/CreationDate").value
        'If pdfDoc.Info.hasKey("/ModDate") Then lblModDate.Caption = pdfDoc.Info.asDictionary("/ModDate").Value
        If pdfDoc.Info.hasKey("/Producer") Then lblProducer.Caption = pdfDoc.Info.asDictionary("/Producer").value
        If pdfDoc.Info.hasKey("/Subject") Then lblSubject.Caption = pdfDoc.Info.asDictionary("/Subject").value
        
        Dim dict As Dictionary
        Set dict = pdfDoc.Info.asDictionary
        Dim v As Variant
        Dim obj As pdfValue
        For Each v In dict.Keys
            lbInfo.AddItem CStr(v) & "=" & dict.item(v).value
        Next v
    End If
End Sub

Private Sub UpdatePages()
    lblPageCount = pdfDoc.pageCount
End Sub

Private Sub UpdateNamedDestinations()
    If pdfDoc.Dests.valueType = PDF_ValueType.PDF_Dictionary Then
        Dim dict As Dictionary: Set dict = pdfDoc.Dests.asDictionary()
        If Not dict Is Nothing Then
            Dim v As Variant
            Dim obj As pdfValue
            For Each v In dict.Keys
                Set obj = dict(v)
                lbNamedDestinations.AddItem CStr(v) & ":" & BytesToString(obj.serialize())
            Next v
            Set v = Nothing
        End If
        Set dict = Nothing
    End If
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
        UpdateNamedDestinations
    End If
End Sub
