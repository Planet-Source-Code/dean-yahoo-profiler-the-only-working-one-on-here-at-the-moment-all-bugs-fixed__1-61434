Attribute VB_Name = "Module1"
Option Explicit


Public Sub FindHTMLTag(ByVal MainStr As String, ByRef s, ByRef l, Optional ByVal ss As Long = 0)
    ' This function finds an HTML tag in a string and puts the start position and the length of the tag in s and l respectively
    ' the HTML tags are the tags that do contain text that is displayed on your browser
    ' however, tags like <script> .... </script> containg data that should be filtered as a whole
    ' for that purpose, use FindSHTMLTag, which will get a tag from start to end
    
    ' MainStr is the string containing HTML source
    ' s is a return value that represents the position where the HTML tag starts
    ' l is the  lenght of the HTML text
    ' ss is the start string position, by default it points to the start position of the string
    ' if no tag is found l will return with a negative value...
    ' to ensure proper operation check the value of l
    
    Dim sl, ll As Long
    
    If IsNull(MainStr) Or Len(MainStr) < 3 Then 'the string passed can not be a tag at all
        s = 0
        l = -1
        Exit Sub
    End If
    
    sl = InStr(ss + 1, MainStr, "<")
    ll = InStr(sl + 1, MainStr, ">")
    
    If sl = 0 Or ll = 0 Then
        s = 0
        l = -2
        Exit Sub
    End If
    
    s = sl - 1
    l = ll - s
    
End Sub

Public Sub FindSHTMLTag(ByVal MainStr As String, ByVal TagStr As String, _
                        ByRef s, ByRef l, Optional ByVal ss As Long = 0)
    ' This function finds tags like <str> .... </str>
    ' and returns the start position for <str> all the way to </str>
    
    ' MainStr is the string containing HTML source
    ' TagStr is sort of the tag type i.e. <script ...>
    ' s is a return value that represents the position where the HTML tag starts
    ' l is the  lenght of the whole text
    ' ss is the start string position, by default it points to the start position of the string
    ' if no tag is found l will return with a negative value...
    ' to ensure proper operation check the value of l
    
    Dim s1, s2, l2 As Long
    Dim str1, str2 As String
    
    If IsNull(MainStr) Or Len(MainStr) < 3 Or _
       IsNull(TagStr) Or Len(TagStr) = 0 Then 'the strings passed can not be a tag at all
        s = 0
        l = -1
        Exit Sub
    End If
    
    str1 = "<" + TagStr
    str2 = "</" + TagStr
    
    s1 = InStr(ss + 1, MainStr, str1, vbTextCompare)
    s2 = InStr(s1 + 1, MainStr, str2, vbTextCompare)
    If s1 = 0 Or s2 = 0 Then
        s = 0
        l = -2
        Exit Sub
    End If
    
    l2 = InStr(s2, MainStr, ">", vbTextCompare)
    If l2 = 0 Or IsNull(l2) Then
        s = 0
        l = -2
        Exit Sub
    End If
    
    
    s = s1 - 1
    l = l2 - s
    
End Sub

Public Function NormalizeTags(ByVal str As String) As String
    ' this function eliminates the spaces from the tag commands
    ' < script ...> will become <script ...>
    ' the main usage is Normalizing all the HTML string before looking for specific tags i.e. <title>
    
    ' It is not a bad idea to get rid of multiple space before normalization
    str = CleanDoubleSpaces(str)
    
    Do While InStr(1, str, "< ") > 0 ' continue as long as there are spaces between "<" and the tag commands
        str = Replace(str, "< ", "<")
    Loop
    
    NormalizeTags = str
End Function

Public Function Convert2Space(ByVal str As String) As String
    ' The main purpose of this function is to convert formatting characters
    ' such as tabs and line breaks to spaces, and then simply by calling the
    ' CleanDoubleSpaces function, the text is even more filtered
    ' Ascii Values 9, 10, and 13 convert to tab, linefeed, and carriage return characters, respectively
    
    Dim cspace, ctab, clinefeed, creturn, cc As String
    
    cspace = Chr(32)
    ctab = Chr(9)
    clinefeed = Chr(10)
    creturn = Chr(13)
    
    ' first replace tabs with space
    str = ReplaceAscii(str, ctab, cspace)
    ' then replace carriage return with linefeed and get rid of double spaces
    str = ReplaceAscii(str, creturn, clinefeed)
    str = CleanDoubleSpaces(str)
    ' then replace a space and linefeed combination with a linefeed
    cc = cspace + clinefeed
    str = ReplaceAscii(str, cc, clinefeed)
    ' then replace a linefeed and space combination with a linefeed
    cc = clinefeed + cspace
    str = ReplaceAscii(str, cc, clinefeed)
    ' then filter duplicated linefeeds
    str = CleanRepeatedAscii(str, 10)
    
    Convert2Space = str
End Function
Public Function Convert2SingleLine(ByVal str As String) As String
    ' This function simply converts the string which possibly may contain
    ' linefeed and carriage returns into a single line string
    Dim cspace, ctab, clinefeed, creturn, cc As String
    
    cspace = Chr(32)
    ctab = Chr(9)
    clinefeed = Chr(10)
    creturn = Chr(13)
    
    ' Replace all the tabs with space
    str = ReplaceAscii(str, ctab, cspace)
    ' Replace all the linefeeds with space
    str = ReplaceAscii(str, clinefeed, cspace)
    ' Replace all the carriage returns with space
    str = ReplaceAscii(str, creturn, cspace)
    ' Finally, get rid of double spaces
    str = CleanDoubleSpaces(str)
    
    If Left(str, 1) = " " Then
        str = Right(str, Len(str) - 1)
    End If
    Convert2SingleLine = str
End Function

Private Function ReplaceAscii(ByVal str As String, ByVal f As String, _
    ByVal r As String) As String
    ' This function simply finds the character represented with f
    ' and replaces it with the character represented with r
    Dim ss As Long

    If IsNull(str) Or IsEmpty(str) Then
        ReplaceAscii = str
        Exit Function
    End If
    
    Do While True
        ss = InStr(str, f)
        If ss > 0 Then
            str = Replace(str, f, r)
        Else
            Exit Do
        End If
    Loop
    ReplaceAscii = str
End Function

Public Function DeleteString(ByVal str As String, ByVal s As Long, ByVal l As Long) As String
    ' This function deletes the part of a string indicated by
    ' s, the starting position
    ' l, the length of the string to be deleted
    If l = 0 Then
        DeleteString = str
        Exit Function
    End If
    DeleteString = Left(str, s) + Right(str, Len(str) - s - l)
End Function

Public Function CleanHTMLTags(ByVal str As String) As String
    ' This function finds and deletes the HTML tags
    ' To clean a tag from start to end i.e. <a html=...> ... </a> use the function CleanSHTMLTags

    Dim s, l As Long 'variables to hold the start position and the length of HTML tags found in the current string
    
    ' Not a bad idea to first normalize the code passed, just in case...
    ' Moron NormalizeTags...
    str = NormalizeTags(str)
    
    If IsNull(str) Then
        str = "no text to clean from HTML tags"
        Exit Function
    End If
    
    Do While l >= 0
        str = DeleteString(str, s, l)
        FindHTMLTag str, s, l
    Loop
    
    str = CleanDoubleSpaces(str)
    CleanHTMLTags = str
End Function

Public Function CleanSHTMLTags(ByVal MainStr As String, ByVal TagStr As String) As String
    ' This function finds and deletes the HTML tags from the begining to the end of the tag including whatever is in between
    ' To clean a tag from start to end i.e. <a html=...> ... </a> use this function

    ' MainStr contains the text to be cleaned
    ' TagStr is the tag type to be cleaned from the text
    
    Dim s, l As Long 'variables to hold the start position and the length of HTML tags found in the current string
    
    
    If IsNull(MainStr) Then
        MainStr = "no text to clean from HTML tags"
        Exit Function
    End If
    
    ' Not a bad idea to first normalize the code passed, just in case...
    ' Moron NormalizeTags...
    MainStr = NormalizeTags(MainStr)
    
    Do While l >= 0
        MainStr = DeleteString(MainStr, s, l)
        FindSHTMLTag MainStr, TagStr, s, l
    Loop
    
    MainStr = CleanDoubleSpaces(MainStr)
    CleanSHTMLTags = MainStr
End Function

Public Function CleanRepeatedAscii(ByVal str As String, ByVal ac As Long) As String
    ' This function simply gets rid of the repeated ascii charactes...
    ' It really helps after filtering an HTML code since lots of line breaks will be left
    ' Also double spaces are eliminated this way
    ' str is the HTML code, and ac is the repeated ascii code to be cleaned
    Dim s, l, ss, ls As Long
    Dim ch, ch2 As String
    
    ch = Chr(ac)
    ch2 = ch + ch
    
    If IsNull(str) Or IsEmpty(str) Then
        str = "no text to clean from specified Ascii code"
        Exit Function
    End If
    
    Do While True
        ss = InStr(str, ch2)
        If ss > 0 Then
            str = Replace(str, ch2, ch)
        Else
            Exit Do
        End If
    Loop
    CleanRepeatedAscii = str
End Function

Public Function CleanDoubleSpaces(ByVal str As String) As String
    ' Mostly after cleaning tags, double and more spaced will remain in the text
    ' this function simply gets rid of them
    
    str = CleanRepeatedAscii(str, "32")

    CleanDoubleSpaces = str
End Function

Public Function HTML2Text(ByVal str As String) As String
    ' This function gets in the source code of an HTML page
    ' and returns the text extracted from the source code
    ' It will first normalize the code (read the NormalizeTags function for further info)
    ' get rid of the head and then SHTML tags
    ' Then get rid of the HTML tags and finally, clean the double spaces
    str = NormalizeTags(str)
    str = CleanSHTMLTags(str, "head")
    str = CleanSHTMLTags(str, "script")
    str = CleanHTMLTags(str)
    str = CleanRepeatedAscii(str, 13)
    str = CleanRepeatedAscii(str, 10)
    str = CleanRepeatedAscii(str, 9)
    str = CleanDoubleSpaces(str)
    HTML2Text = str
End Function


Public Function GetStringBetween(ByVal str As String, ByVal str1 As String, ByVal str2 As String, _
                                 Optional ByVal st As Long = 0) As String
    ' This function gets in a string and two keywords
    ' and returns the string between the keywords
    
    Dim s1, s2, s, l As Long
    Dim foundstr As String
    
    s1 = InStr(st + 1, str, str1, vbTextCompare)
    s2 = InStr(s1 + 1, str, str2, vbTextCompare)
    
    If s1 = 0 Or s2 = 0 Or IsNull(s1) Or IsNull(s2) Then
        foundstr = str
    Else
        s = s1 + Len(str1)
        l = s2 - s
        foundstr = Mid(str, s, l)
    End If
    
    GetStringBetween = foundstr
End Function

Public Function GetTitle(ByVal str As String) As String
    ' this function returns the title of an HTML page
    Dim title As String
    
    ' Not a bad idea to first normalize the code passed, just in case...
    ' Moron NormalizeTags...
    str = NormalizeTags(str)
    
    title = GetStringBetween(str, "<title>", "</title>")
    If title = str Then
        GetTitle = "No Title Detected"
    Else
        GetTitle = title
    End If
End Function





