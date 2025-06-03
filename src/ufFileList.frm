VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufFileList 
   Caption         =   "Combine PDF files"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8250.001
   OleObjectBlob   =   "ufFileList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public list As Variant


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

Private Sub spUpDown_SpinDown()
    swapOrder lbFiles.ListIndex, lbFiles.ListIndex + 1
    If lbFiles.ListIndex < (lbFiles.ListCount - 1) Then lbFiles.ListIndex = lbFiles.ListIndex + 1
End Sub

Private Sub spUpDown_SpinUp()
    swapOrder lbFiles.ListIndex, lbFiles.ListIndex - 1
    If lbFiles.ListIndex > 0 Then lbFiles.ListIndex = lbFiles.ListIndex - 1
End Sub

Private Sub UserForm_Activate()
    Dim i As Long
    For i = LBound(list) To UBound(list)
        lbFiles.AddItem list(i)
    Next
End Sub

Private Sub UserForm_Initialize()
    Dim testData() As String
    testData = Split("file1,file2,file3,file4,file5", ",")
    list = testData
End Sub

