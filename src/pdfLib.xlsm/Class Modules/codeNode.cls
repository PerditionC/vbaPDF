VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "codeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Type codeNode
Option Explicit

Public extraBitsLen As Integer    ' how many extra bits to extract and add to min value
Public minValue As Long           ' starting value for the corresponding code this node represents
Public maxValue As Long           ' max value for the code value this node represents

