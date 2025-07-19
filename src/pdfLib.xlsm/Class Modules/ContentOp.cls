VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContentOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'describes one PDF operator after tokenisation
' class instead of a Type to enable storing in Collection and Dictionary as a COM object
Option Explicit

Public opType   As PdfTextOp  ' enum indicating op
Public opName   As String     ' alternate, text name of op
Public operands As Dictionary ' of pdfValue

Private Sub Class_Initialize()
    'opType = PdfTextOp.opUnknown
    'opName = vbNullString
    Set operands = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set operands = Nothing
End Sub
