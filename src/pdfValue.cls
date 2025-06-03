VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pdfValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' stores a PDF value type
Option Explicit

Public id As Long
Public generation As Long
Public valueType As PDF_ValueType
Public Value As Variant

