### vbaPDF Developer Guide  

---

## 1  Overview  

`vbaPDF` is a pure-VBA library for **reading, writing and merging PDF files** from Excel (untested but should also work with Word or any Office host that supports VBA.)  The design is pdf "value" and "object-level" based – PDF objects are parsed into wrapper classes (`pdfDocument`, `pdfValue`, `pdfStream`, …) so you can inspect or modify the underlying structure without external DLLs.

Main entry points  

| Module/Class | Purpose | Notes |
|--------------|---------|-------|
| **pdfDocument** (class) | Open, parse, edit, save a PDF. | Core API – primary API for interacting with PDFs |
| **Main.bas** | Ready-made helpers: file-picker UserForms and `CombinePDFs`. | Reference examples |
| **pdfValue** | Universal wrapper for every PDF value type. | Internal, represents all values stored in a PDF |
| **pdfStream** | Handles `stream … endstream` objects, incl. deflate decode. | Internal, not usually directly used |
| **UserForms** `ufPdfInfo`, `ufFileList`, `ufBookmarkEditor` | UI widgets for demos (optional). | GUI for reference examples |

See [API.md](API.md) for more details.

---

## 2  Feature Matrix  

| Feature | Status | Where / Comments |
|---------|--------|------------------|
| Load PDF (incl. 1.0 – 2.0 headers, xref tables & xref streams) | ✅ Implemented | `pdfDocument.loadPdf` (only subset of PDF specification supported) |
| Parse full object tree into memory | ⚠️ | `pdfDocument.parsePdf` (implemented for all supported features, may have issues with unsupported features) |
| Save PDF (re-writes xref, trailer) | ✅ | `savePdf / savePdfAs` |
| Append / remove pages | ✅ | `AddPages`, `RemovePage` (example below) |
| Combine/merge PDFs | ✅ | `Main.CombinePDFs` |
| Read/write Info dictionary (/Title, /Author …) | ✅ | `pdfDocument.Info` |
| Read/write Catalog, /Pages, /Outlines, /Dests | ✅ | dedicated helpers |
| Named destinations | ✅ | `AddNamedDestinations` |
| Bookmarks (Outlines) read + basic write | ⚠️ Partial | Writing works but deep nesting helpers TODO |
| Text extraction | ⚠️ Partial | In progress building content-stream parser, will **NOT** do OCR or advanced extraction |
| Form fields (AcroForm) | 🚧 TODO | possibly in the future if needed or requested |
| Encryption / passwords | ❌ Not planned |
| Incremental update | ❌ Not planned (library always rewrites full file) |

---

## 3  Quick-start  

```vba
'--- add modules ---
' Import every *.bas / *.cls / *.frm file from /src into your VBA project.
' No references or external libraries required.

Sub HelloPDF()
    Dim pdf As pdfDocument: set pdf = New pdfDocument
    pdf.OpenPdf "C:\Docs\Test.pdf"      ' Load
    MsgBox "Pages: " & pdf.PageCount    ' Inspect
    pdf.SavePdf "C:\Docs\Test-copy.pdf"    ' Save copy
End Sub
```

---

## 4  Cookbook  

### 4.1 Open, inspect & close  

```vba
Dim doc As pdfDocument
set doc = pdfDocument.pdfDocument("C:\tmp\report.pdf")

Debug.Print "Title    : "; doc.Title
Debug.Print "Producer : "; doc.Info.asDictionary("/Producer").Value
Debug.Print "Pages    : "; doc.PageCount

set doc = Nothing
```

---

### 4.2 Iterate pages & extract text (once implemented)

```vba
' CURRENTLY NOT IMPLEMENTED – placeholder API
Dim i As Long
For i = 1 To doc.PageCount
    Debug.Print doc.Pages(i).ExtractText   ' Coming soon
Next i
```

---

### 4.3 Add or remove pages  

```vba
'--- add a blank page to the end ---
doc.AddPage        ' Wrapper around NewPage + AddPages
```

```vba
' CURRENTLY NOT IMPLEMENTED – placeholder API
'--- remove third page (0-based index) ---
doc.RemovePage 2
```

---

### 4.4 Named Destinations  

```vba
'Create a destination that shows page 1 at Fit Zoom
Dim dest As pdfValue
Set dest = doc.NewDestination(0, PDF_FIT.PDF_FIT)   ' zero-based page index

Dim name As pdfValue
Set name = pdfValue.NewNameValue("/IntroPage")

doc.AddNamedDestinations name, dest
```

---

### 4.5 Bookmarks / Outlines (basic)  

```vba
'Top-level Outlines node (creates one if missing)
doc.AddOutlines

'Child outline linking to page 1
Dim oItem As pdfValue
Dim defaults As New Dictionary
defaults("/Title") = pdfValue.NewValue("Introduction")
defaults("/Dest")  = dest           ' dest from previous snippet

Set oItem = doc.NewOutlineItem(Nothing, defaults)
doc.AddOutlineItem doc.Outlines, oItem
```

---

### 4.6 Combine multiple PDFs (ready-made)  

```vba
Sub Merge()
    Dim files() As String
    files = Array("C:\a.pdf", "C:\b.pdf", "C:\c.pdf")
    CombinePDFs files, "C:\out\merged.pdf"
End Sub
```

`Main.bas` also contains `PickAndCombinePdfFiles` which displays file-picker/UI forms so end-users can do the same without writing code.

---

### 4.7 Renumber Object IDs (advanced)  

Useful when you splice objects from one document into another.

```vba
Dim base As Long
base = targetDoc.nextObjId          ' first free id in destination, note it auto-advances to next call =+1
sourceDoc.renumberIds base          ' shifts every object id
```

---

## 5  UserForm helpers  

| Form | What it does | How to launch |
|------|--------------|---------------|
| `ufPdfInfo` | View metadata, page count & named destinations of a single PDF. | ```VBA: frm.Show``` – if no `pdfDocument` is pre-assigned it prompts for a file. |
| `ufFileList` | Lets end-user re-order an array of filenames before merge. | Used inside `Main.PickAndCombinePdfFiles`. |
| `ufBookmarkEditor` | Prototype bookmark editor. | Experimental. |

---

## 6  Potential future extensions to the API  

1. **Text content stream parser**  
   • Walk graphics operators, recognise `TJ`, `Tj`, etc. to build a `Page.ExtractText()` API.  
2. **Form fields (AcroForm)**  
   • Add `pdfDocument.Forms` returning a collection of field objects with `.Value` property for read/write.  
3. **Stream filters**  
   • Currently supports `/FlateDecode` (deflate); add `/LZWDecode`, `/ASCII85Decode`, etc.  

---

## 7  Limitations  

* Entire file loaded into memory – not suitable for very large PDFs.  
* No incremental-update; save rewrites the full document.  
* No encryption/signature support.  
* Spec compliance for features supported (file a bug if non-compliance found): tested mainly with PDF 1.4–1.7, and 2.0 produced by Microsoft Print to PDF and Acrobat Save As of existing documents.

---

## 8  Minimal cheat-sheet  

```vba
Dim pdf As pdfDocument
set pdf = New pdfDocument       ' create
pdf.OpenPdf "in.pdf"            ' load

Debug.Print pdf.PageCount       ' -> number
pdf.Title = "New title"         ' write /Info

pdf.RemovePage 1                ' drop first page
pdf.SavePdfAs "out.pdf"            ' write to disk
```

---

Contributions are welcome – see the `Issues` tab on GitHub for known issues.
