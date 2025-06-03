Attribute VB_Name = "fileManip"
' basic file manipulation
Option Explicit

' reads in contents of filename returning as a Byte() array
' fileLen is set to size in bytes of the file
Function readFile(ByVal fileName As String, ByRef fileLen As Long) As Byte()
    On Error GoTo errHandler
    Dim fileNum As Integer
    Dim content() As Byte
    
    ' Open file and read content
    fileNum = FreeFile
    Open fileName For Binary Access Read Shared As #fileNum
    fileLen = LOF(fileNum)
    ReDim content(fileLen - 1)
    Get #fileNum, , content
    
cleanup:
    On Error Resume Next
    Close #fileNum
    readFile = content
    Exit Function
errHandler:
    Debug.Print "Error: " & Err.Description & " (" & Err.Number & ")"
    fileLen = 0
    ReDim content(0)
    Resume cleanup
End Function

