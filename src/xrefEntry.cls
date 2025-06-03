VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xrefEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' catalog entry (entry in xref table)
Option Explicit

Public id As Long
Public offset As Long       ' 0 for free objects, offset in bytes from beginning of file for regular objects, index into object for embedObjs
Public generation As Long
Public nextFreeId As Long   ' 0 for live objects
Public isFree As Boolean    ' "n"=live/valid, "f"=free/not in use

' only used if object is part of an object stream
Public isEmbeded As Boolean ' True then next fields valid, False they are unused
Public embedObjId As Long   ' which (the id) stream object stream this obj is embedded in, generation implicitly 0
