# vbaPDF API – Future Content Stream Text Extraction

---

## Public API

### `Function GetPageText(PageIndex As Long) As String`
Extracts all visible text from a single page, returning a plain Unicode string. This is the most common entry point for users who want the text content of a page.

**Parameters:**
- `PageIndex`: The zero-based index of the page to extract text from.

**Returns:**
- A `String` containing all text from the specified page, in logical reading order.

**Example:**
```vba
Dim txt As String
Set pdf = New pdfDocument
pdf.LoadFromFile "sample.pdf"
txt = pdf.GetPageText(0)
MsgBox txt
```

---

### `Function GetDocumentText() As String`
Extracts all text from the entire document, concatenating the text from each page in order.

**Returns:**
- A `String` containing all text from the document, with page breaks represented by `vbFormFeed` (Chr(12)).

**Example:**
```vba
Dim allText As String
allText = pdf.GetDocumentText()
Debug.Print allText
```

---

### `Function GetPageTextWithLayout(PageIndex As Long, Optional AsHtml As Boolean = False) As Variant`
Extracts text from a page, returning a collection of text fragments with position and font information. Optionally, returns HTML representing the page's text layout and basic formatting.

**Parameters:**
- `PageIndex`: The zero-based index of the page.
- `AsHtml`: If `True`, returns a string of HTML; if `False` (default), returns a `Collection` of text fragments.

**Returns:**
- If `AsHtml = False`: A `Collection` of `TextFragment` objects (see below).
- If `AsHtml = True`: A `String` containing HTML markup for the page's text.

**Example:**
```vba
Dim html As String
html = pdf.GetPageTextWithLayout(0, True)
Debug.Print html
```

---

### `Sub SetFontUnicodeMap(FontName As String, Map As Dictionary)`
Sets a mapping from font-encoded character codes to Unicode values for a given font. This is necessary for subset or custom-encoded fonts.

**Parameters:**
- `FontName`: The name or resource identifier of the font.
- `Map`: A `Scripting.Dictionary` mapping font codes (as `Long` or `String`) to Unicode code points (as `String`).

**Example:**
```vba
Dim map As Object
Set map = CreateObject("Scripting.Dictionary")
map(65) = "A"
map(66) = "B"
pdf.SetFontUnicodeMap "F1", map
```

---

## Low-Level API

### `Function ParseContentStream(PageIndex As Long) As Collection`
Parses the raw content stream of a page, returning a collection of low-level text drawing operations (e.g., TJ, Tj, Td, Tf, etc.).

**Parameters:**
- `PageIndex`: The zero-based index of the page.

**Returns:**
- A `Collection` of `ContentOp` objects, each representing a PDF drawing/text operation with operands.

---

### `Function MapFontCodeToUnicode(FontName As String, Code As Variant) As String`
Translates a font-encoded character code to its Unicode equivalent using the current mapping for the font.

**Parameters:**
- `FontName`: The font resource name.
- `Code`: The code from the content stream (as `Long` or `String`).

**Returns:**
- The Unicode string for the code, or `?` if unmapped.

---

## Internal/Private Helpers

### `Function ExtractTextFragments(Ops As Collection, FontMaps As Dictionary) As Collection`
Given a collection of parsed content stream operations and font mappings, returns a collection of `TextFragment` objects with text, position, and font info.

---

### `Function BuildHtmlFromFragments(Fragments As Collection) As String`
Converts a collection of `TextFragment` objects into an HTML string, preserving basic layout and formatting (e.g., lines, spaces, font size, bold/italic if available).

---

### `Type TextFragment`
A structure representing a piece of text with its position and style.
- `Text As String`
- `X As Double`
- `Y As Double`
- `FontName As String`
- `FontSize As Double`
- `IsBold As Boolean`
- `IsItalic As Boolean`

---

### `Type ContentOp`
A structure representing a single PDF content stream operation.
- `OpName As String` (e.g., "TJ", "Tj", "Td", "Tf")
- `Operands As Variant`

---

## Additional Implementation Notes (Remove after implementation)

- Parsing should respect the PDF specification for text extraction, including handling of text state, positioning, and font selection.
- For HTML output, follow the PDF 2.0 spec for logical structure where possible, but fall back to spatial layout if structure is not available.
- For subset fonts, require user to provide a mapping or attempt to auto-detect if ToUnicode CMap is present in the PDF.
- The `TextFragment` collection should be sorted in reading order (left-to-right, top-to-bottom for most Western PDFs).
- For simple use cases, `GetPageText` and `GetDocumentText` should be fast and require no font mapping unless subset fonts are used.
- No OCR or AI-based recognition is performed; only text that is present in the content stream is extracted.
- Consider providing a utility to dump all font names and their encoding types for user inspection.

---

> **End of Text Extraction API Section**
Here is the future-facing API design for content stream text extraction in `vbaPDF`, structured for insertion into your documentation after the existing functions. This is available as a downloadable file named `API-future.md` in the Code playground.

---

## [Section: Text Extraction API]

> **Insert after existing documented functions.**  
> (e.g., after `<!-- TEXT EXTRACTION API START -->` or similar marker.)

---

## Public API

### `Function GetPageText(PageIndex As Long) As String`
Extracts all visible text from a single page, returning a plain Unicode string. This is the most common entry point for users who want the text content of a page.

**Parameters:**  
- `PageIndex`: The zero-based index of the page to extract text from.

**Returns:**  
- A `String` containing all text from the specified page, in logical reading order.

**Example:**
```vba
Dim txt As String
Set pdf = New pdfDocument
pdf.LoadFromFile "sample.pdf"
txt = pdf.GetPageText(0)
MsgBox txt
```

---

### `Function GetDocumentText() As String`
Extracts all text from the entire document, concatenating the text from each page in order.

**Returns:**  
- A `String` containing all text from the document, with page breaks represented by `vbFormFeed` (Chr(12)).

**Example:**
```vba
Dim allText As String
allText = pdf.GetDocumentText()
Debug.Print allText
```

---

### `Function GetPageTextWithLayout(PageIndex As Long, Optional AsHtml As Boolean = False) As Variant`
Extracts text from a page, returning a collection of text fragments with position and font information. Optionally, returns HTML representing the page's text layout and basic formatting.

**Parameters:**  
- `PageIndex`: The zero-based index of the page.
- `AsHtml`: If `True`, returns a string of HTML; if `False` (default), returns a `Collection` of text fragments.

**Returns:**  
- If `AsHtml = False`: A `Collection` of `TextFragment` objects (see below).
- If `AsHtml = True`: A `String` containing HTML markup for the page's text.

**Example:**
```vba
Dim html As String
html = pdf.GetPageTextWithLayout(0, True)
Debug.Print html
```

---

### `Sub SetFontUnicodeMap(FontName As String, Map As Dictionary)`
Sets a mapping from font-encoded character codes to Unicode values for a given font. This is necessary for subset or custom-encoded fonts.

**Parameters:**  
- `FontName`: The name or resource identifier of the font.
- `Map`: A `Scripting.Dictionary` mapping font codes (as `Long` or `String`) to Unicode code points (as `String`).

**Example:**
```vba
Dim map As Object
Set map = CreateObject("Scripting.Dictionary")
map(65) = "A"
map(66) = "B"
pdf.SetFontUnicodeMap "F1", map
```

---

## Low-Level API

### `Function ParseContentStream(PageIndex As Long) As Collection`
Parses the raw content stream of a page, returning a collection of low-level text drawing operations (e.g., TJ, Tj, Td, Tf, etc.).

**Parameters:**  
- `PageIndex`: The zero-based index of the page.

**Returns:**  
- A `Collection` of `ContentOp` objects, each representing a PDF drawing/text operation with operands.

---

### `Function MapFontCodeToUnicode(FontName As String, Code As Variant) As String`
Translates a font-encoded character code to its Unicode equivalent using the current mapping for the font.

**Parameters:**  
- `FontName`: The font resource name.
- `Code`: The code from the content stream (as `Long` or `String`).

**Returns:**  
- The Unicode string for the code, or `?` if unmapped.

---

## Internal/Private Helpers

### `Function ExtractTextFragments(Ops As Collection, FontMaps As Dictionary) As Collection`
Given a collection of parsed content stream operations and font mappings, returns a collection of `TextFragment` objects with text, position, and font info.

---

### `Function BuildHtmlFromFragments(Fragments As Collection) As String`
Converts a collection of `TextFragment` objects into an HTML string, preserving basic layout and formatting (e.g., lines, spaces, font size, bold/italic if available).

---

### `Type TextFragment`
A structure representing a piece of text with its position and style.
- `Text As String`
- `X As Double`
- `Y As Double`
- `FontName As String`
- `FontSize As Double`
- `IsBold As Boolean`
- `IsItalic As Boolean`

---

### `Type ContentOp`
A structure representing a single PDF content stream operation.
- `OpName As String` (e.g., "TJ", "Tj", "Td", "Tf")
- `Operands As Variant`

---

## Additional Implementation Notes (Remove after implementation)

- Parsing should respect the PDF specification for text extraction, including handling of text state, positioning, and font selection.
- For HTML output, follow the PDF 2.0 spec for logical structure where possible, but fall back to spatial layout if structure is not available.
- For subset fonts, require user to provide a mapping or attempt to auto-detect if ToUnicode CMap is present in the PDF.
- The `TextFragment` collection should be sorted in reading order (left-to-right, top-to-bottom for most Western PDFs).
- For simple use cases, `GetPageText` and `GetDocumentText` should be fast and require no font mapping unless subset fonts are used.
- No OCR or AI-based recognition is performed; only text that is present in the content stream is extracted.
- Consider providing a utility to dump all font names and their encoding types for user inspection.

---

> **End of Text Extraction API Section**

---
