Attribute VB_Name = "LibStringManip"
' basic String manipulations
Option Explicit

#Const FASTCOMPARE_ = True

#If FASTCOMPARE_ Then
'Windows function (unicode), to use call with lstrcmpi(StrPtr(str1),StrPtr(str2))
Public Declare PtrSafe Function lstrcmpi Lib "kernel32" Alias "lstrcmpiW" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As Long
#End If


' Function: IsBlank
' Determins if a String contains any non-whitespace characters.
'
' Parameters:
'   str - the String to examine, not modified
'
' Returns:
'   True - if *str* is a zero-length string "" after removing
'       leading and trailing whitespace
'   False - if *str* contains any non-whitespace characters
'
' See Also:
'   <IsBlankorNull>, <IsWhiteSpace>, <TrimWS>, <IsMatch>
'
Function IsBlank(ByVal str As String) As Boolean
    ' remove leading & trailing whitespace, then determine if zero-length string ""
    IsBlank = LenB(TrimWS(str)) = 0
End Function


' Function: IsBlankOrNull
' Determins if a String contains any non-whitespace characters or IsNull.
'
' Parameters:
'   str - the String to examine, not modified
'
' Returns:
'   True - if *str* is a zero-length string "" after removing
'       leading and trailing whitespace
'       or if IsNull(str) returns True
'   False - if *str* contains any non-whitespace characters
'
' See Also:
'   <IsBlank>, <IsWhiteSpace>, <TrimWS>, <IsMatch>
'
Function IsBlankOrNull(ByVal str As Variant) As Boolean
    ' first check if is Null
    If IsNull(str) Then
        IsBlankOrNull = True
    Else
        ' remove leading & trailing whitespace, then determine if zero-length string ""
        IsBlankOrNull = LenB(TrimWS(str)) = 0
    End If
End Function


' Function: IsWhiteSpace
' Determines if the first character of a String is a whitespace character.
' Whitespace includes space (0x20, " "), [horizontal] tab (0x09, \t), form feed (0x0c, \f)
'   carriage return (0x0d, \r), and linefeed (0x0a, \n)
'
' Paramters:
'   str - the String to examine, not modified
'
' Returns:
'   True - if the first letter in *str* is any whitespace character
'   False - if teh first letter in *str* is not a whitespace character or
'       *str* is a zero-length string ""
'
' See Also:
'   <IsBlank>, <IsBlankOrNull>, <TrimWS>, <IsMatch>
'
' Note:
' textual space includes horizontal tab (0x09), line feed (0x0a), vertical tab (0x0b),
' form feed (0x0c), carriage return (0x0d), space (0x20), nonbreaking space (0xa0),
' en space (0x2002), em space (0x2003), figure space (0x2007), punctuation space (0x2008),
' thin space (0x2009), hair space (0x200a), zero width space (0x200b), and ideographic space (0x300)
Function IsWhiteSpace(ByVal str As String) As Boolean
    ' assume not whitespace until proven otherwise (by default False is returned)
    ' a zero-length string "" is not considered as whitespace
    If LenB(str) <> 0 Then
        ' get character code of str and evaluate type of character it is
        Select Case AscW(str)
            ' is it a whitespace character
            Case 32, 10, 13, 9, 19  ' space, lf, cr, tab, and form-feed (other textual spaces are not valid for PDF whitespace)
                IsWhiteSpace = True
        End Select
    End If
End Function


' Function: TrimWS
' Removes all leading and all trailing whitespace characters from a String.
' Any whitespace between non-whitespace characters is preserved.
'
' Paramters:
'   str - the String, not modified
'
' Returns:
'   a copy of *str* with leading and trailing whitespace removed
'
' See Also:
'   <IsBlank>, <IsBlankOrNull>, <IsWhiteSpace>, <IsMatch>
'
Function TrimWS(ByVal str As String) As String
    ' Note: do not use this as it fails to remove new lines and
    ' also changes whitespace between non-whitespace characters to at most 1 space
    ' TrimWS = Application.WorksheetFunction.Trim(str)
    
    Dim ndx As Long
    
    ' use builtin to remove initial leading & trailing spaces
    str = Trim$(str)
    
    ' only proceed not a blank string
    If LenB(str) > 0 Then
#If False Then
        ' if copying from a bulletted list, usually the bullet get added but is unwnated, so strip it too
        ndx = Asc$(Left$(str, 1))
        If (i = 149) Or (i = 160) Then ' 149=bullet point, 160=non-breaking space
            str = Trim$(Mid$(str, 2))
        End If
#End If

        ' start at first character in string
        ndx = 1
        
        ' while there are still characters in string, advance index to first non-whitespace character index
        Do While IsWhiteSpace(Mid(str, ndx, 1))
            ndx = ndx + 1
            DoEvents
        Loop
        
        ' remove whitespace characters from beginning of the string
        str = Mid(str, ndx)
        
        ' start at the last character in string
        ndx = Len(str)
        
        ' if still non-whitespace characters, then see if any trailing whitespace
        Do While ndx > 0
            If Not IsWhiteSpace(Mid$(str, ndx, 1)) Then Exit Do
            ndx = ndx - 1
            DoEvents
        Loop
        If ndx > 0 Then
            str = Left$(str, ndx)
        Else
            str = vbNullString
        End If
    End If
    
    TrimWS = str
End Function


' Function: IsMatch
' Compares two strings to determine if same value after removing
' all leading and all trailing whitespace and ignoring case.
'
' Parameters:
'   str1 - the Strings to compare, not modified
'   str2 - the Strings to compare, not modified
'
' Returns:
'   True - if *str1* = *str2* ignoring case and extra whitespace
'   False - if *str1* <> *str2* ignoring case and extra whitespace or
'       Len(*str1*) <> Len(*str2*) after removing extra whitespace
'
' See Also:
'   <IsBlank>, <IsBlankOrNull>, <IsWhiteSpace>, <TrimWS>
'
Function IsMatch(ByVal str1 As String, ByVal str2 As String) As Boolean
    ' 1st remove any extra whitespace
    str1 = TrimWS(str1)
    str2 = TrimWS(str2)
    
#If FASTCOMPARE_ Then
    ' call the Windows API (Unicode version) to do case insensitive compare
    IsMatch = lstrcmpi(StrPtr(str1), StrPtr(str2)) = 0
#Else
    ' standard VB string compare, case insensitive
    IsMatch = StrComp(str1, str2, vbTextCompare) = 0
#End If
End Function

