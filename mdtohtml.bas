DECLARE FUNCTION TextHeading$ (myString$, mdTag$, htmlStart$, htmlEnd$)
DECLARE FUNCTION TextMarkup$ (myString$, mdTag$, htmlStart$, htmlEnd$)
DECLARE FUNCTION FormatLinks$ (myString$)
DECLARE FUNCTION AddToHTML$ (html$, update$)

If Command$ = "" Then
    Input "Markdown file: ", fileMD$
    Input "CSS Theme:", fileCSS$
ElseIf InStr(Command$, " ") > 0 Then
    fileMD$ = _StartDir$ + "\" + Left$(Command$, InStr(Command$, " ") - 1)
    fileCSS$ = Right$(Command$, Len(Command$) - Len(Left$(Command$, InStr(Command$, " "))))
Else
    fileMD$ = _StartDir$ + "\" + Command$
End If
Print fileMD$
Print fileCSS$ + ".css"

'HTML file created for output
fileHTML$ = Left$(fileMD$, InStr(fileMD$, ".") - 1) + ".html"
'fileHTML$ = Left$(fileMD$, InStr(fileMD$, ".") - 1) + ".htm" 'use this line instead when porting to older QB (DOS)

html$ = "<!doctype html>" + Chr$(10)
html$ = html$ + "" + Chr$(10)
html$ = html$ + "<html lang=" + Chr$(34) + "en" + Chr$(34) + ">" + Chr$(10)
html$ = html$ + "<head>" + Chr$(10)
html$ = html$ + "  <meta charset=" + Chr$(34) + "utf-8" + Chr$(34) + ">" + Chr$(10)
html$ = html$ + "  <meta name=" + Chr$(34) + "viewport" + Chr$(34) + " content=" + Chr$(34) + "width=device-width, initial-scale=1" + Chr$(34) + ">" + Chr$(10)
html$ = html$ + "  <title></title>" + Chr$(10)
html$ = html$ + "  <link rel=" + Chr$(34) + "stylesheet" + Chr$(34) + " href=" + Chr$(34) + "themes/" + fileCSS$ + ".css" + Chr$(34) + ">" + Chr$(10)
html$ = html$ + "</head>" + Chr$(10)
html$ = html$ + "" + Chr$(10)
html$ = html$ + "<body>" + Chr$(10)

InUnorderedList = 0
InOrderedList = 0
InBlockQuote = 0

Open fileMD$ For Input As #1
Open fileHTML$ For Output As #2

While Not EOF(1)
    Line Input #1, markdown$ 'Loads (next) line from MD file

    'Check if still in a list
    If InUnorderedList = 1 And Left$(markdown$, 2) <> "* " And Left$(markdown$, 2) <> "- " And Left$(markdown$, 2) <> "+ " Then
        html$ = AddToHTML$(html$, "</ul>")
        InUnorderedList = 0
    End If
    IF InOrderedList = 1 AND LEFT$(markdown$, 3) <> "1. " AND LEFT$(markdown$, 3) <> "2. " AND LEFT$(markdown$, 3) <> "3. " AND LEFT$(markdown$, 3) <> "4. " AND LEFT$(markdown$, 3) <> "5. " AND LEFT$(markdown$, 3) <> "6. " AND LEFT$(markdown$, 3) <> _
 "7. " AND LEFT$(markdown$, 3) <> "8. " AND LEFT$(markdown$, 3) <> "9. " THEN
        html$ = AddToHTML$(html$, "</ol>")
        InOrderedList = 0
    End If

    'Check for headings
    For i = 1 To 6
        markdown$ = TextHeading$(markdown$, String$(i, "#"), "<h" + LTrim$(Str$(i)) + ">", "</h" + LTrim$(Str$(i)) + ">")
    Next

    'Check for unordered lists
    If Left$(markdown$, 2) = "* " Or Left$(markdown$, 2) = "- " Or Left$(markdown$, 2) = "+ " Then
        If InUnorderedList = 0 Then
            InUnorderedList = 1
            markdown$ = "<ul><li>" + Right$(markdown$, Len(markdown$) - 2) + "</li>"
        Else
            markdown$ = " <li>" + Right$(markdown$, Len(markdown$) - 2) + "</li>"
        End If
        'Check for ordered lists
    ELSEIF LEFT$(markdown$, 3) = "1. " OR LEFT$(markdown$, 3) = "2. " OR LEFT$(markdown$, 3) = "3. " OR LEFT$(markdown$, 3) = "4. " OR LEFT$(markdown$, 3) = "5. " OR LEFT$(markdown$, 3) = "6. " OR LEFT$(markdown$, 3) = "7. " OR LEFT$(markdown$, 3) = _
 "8. " OR LEFT$(markdown$, 3) = "9. " THEN
        If InOrderedList = 0 Then
            InOrderedList = 1
            markdown$ = "<ol><li>" + Right$(markdown$, Len(markdown$) - 2) + "</li>"
        Else
            markdown$ = " <li>" + Right$(markdown$, Len(markdown$) - 2) + "</li>"
        End If
    End If

    'Check for blockquotes
    If Left$(markdown$, 2) = "> " Then
        markdown$ = "<blockquote>" + Right$(markdown$, Len(markdown$) - 2) + "</blockquote>"
    End If

    'Check for bold/italic parts
    markdown$ = TextMarkup$(markdown$, "***", "<em><strong>", "</strong></em>")
    markdown$ = TextMarkup$(markdown$, "**", "<strong>", "</strong>")
    markdown$ = TextMarkup$(markdown$, "*", "<em>", "</em>")
    markdown$ = TextMarkup$(markdown$, "___", "<em><strong>", "</strong></em>")
    markdown$ = TextMarkup$(markdown$, "__", "<strong>", "</strong>")
    markdown$ = TextMarkup$(markdown$, "_", "<em>", "</em>")

    'Check for line breaks
    If Right$(markdown$, 2) = "  " Then
        markdown$ = markdown$ + "<br />"
    End If

    'Check for links
    markdown$ = FormatLinks$(markdown$)

    html$ = AddToHTML$(html$, markdown$)
Wend

html$ = AddToHTML$(html$, "</body>")
html$ = AddToHTML$(html$, "</html>")

Print #2, html$
Close #1
Close #2

System 'Exit program without "Press any key" prompt

Function FormatLinks$ (myString$)
    FormatLinks$ = myString$
    If InStr(myString$, "[") > 0 Then
        linkBefore$ = Left$(myString$, InStr(myString$, "[") - 1)
        linkName$ = Mid$(myString$, InStr(myString$, "[") + 1, InStr(myString$, "]") - InStr(myString$, "[") - 1)
        linkURL$ = Mid$(myString$, InStr(myString$, "(") + 1, InStr(myString$, ")") - InStr(myString$, "(") - 1)
        linkAfter$ = Right$(myString$, Len(myString$) - InStr(myString$, ")"))
        combinedOutput$ = linkBefore$ + "<a href=" + Chr$(34) + linkURL$ + Chr$(34) + ">" + linkName$ + "</a>" + linkAfter$
        FormatLinks$ = FormatLinks$(combinedOutput$)
    End If
End Function

' This function replaces
Function TextHeading$ (myString$, mdTag$, htmlStart$, htmlEnd$)
    TextHeading$ = myString$
    If Left$(myString$, Len(mdTag$) + 1) = mdTag$ + " " Then
        TextHeading$ = htmlStart$ + Right$(myString$, Len(myString$) - Len(mdTag$) - 1) + htmlEnd$
    End If
End Function

' Deals
Function TextMarkup$ (myString$, mdTag$, htmlStart$, htmlEnd$)
    TextMarkup$ = myString$
    If InStr(myString$, mdTag$) > 0 Then
        beforeMarkup$ = Left$(myString$, InStr(myString$, mdTag$) - 1)
        afterMarkup$ = Right$(myString$, Len(myString$) - Len(beforeMarkup$) - Len(mdTag$))
        duringMarkup$ = Left$(afterMarkup$, InStr(afterMarkup$, mdTag$) - 1)
        afterMarkup$ = Right$(afterMarkup$, Len(afterMarkup$) - Len(duringMarkup$) - Len(mdTag$))
        combinedOutput$ = beforeMarkup$ + htmlStart$ + duringMarkup$ + htmlEnd$ + afterMarkup$
        TextMarkup$ = TextMarkup$(combinedOutput$, mdTag$, htmlStart$, htmlEnd$)
    End If
End Function

Function AddToHTML$ (html$, update$)
    AddToHTML$ = html$ + update$ + Chr$(10)
End Function
