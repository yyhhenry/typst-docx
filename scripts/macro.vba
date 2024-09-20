Sub PasteTypstDocx()
    Dim http As Object
    Dim url As String
    Dim result As String
    Dim timestamp As String

    Shell "typst-docx.exe", vbMinimizedNoFocus
    
    If Selection.Type <> wdSelectionIP Then
        Call CutWithoutFinalLinekBreak
    End If
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    timestamp = Format(Now, "yyyymmddHHmmss")

    url = "http://127.0.0.1:5180"
    url = url & "?t=" & timestamp
    
    On Error Resume Next
    http.Open "GET", url, False
    http.Send
    result = http.responseText
    On Error GoTo 0

    If http.Status <> 200 Then
        MsgBox "HTTP Error: " & http.Status & " - " & http.StatusText
        Exit Sub
    End If

    If LCase(Left(result, 5)) = "error" Then
        MsgBox result
    ElseIf result <> "" Then
        originalStyle = Selection.Style
        currentPosition = Selection.Start
        currentLine = Selection.Information(wdFirstCharacterLineNumber)
        
        Selection.InsertFile FileName:=result
        Call DeleteBackToLineEnd

        typstEnd = Selection.Start
        typstEndLine = Selection.Information(wdFirstCharacterLineNumber)
        If currentLine <> typstEndLine Then
            Selection.Start = currentPosition
            Selection.End = typstEnd
        End If
        Selection.Style = originalStyle
    End If
    
    Set http = Nothing
End Sub

Sub DeleteBackToLineEnd()
    Dim currentLine As Integer
    Dim prevLineEnd As Integer
    Dim currentPosition As Long
    Dim lineStart As Long
    Dim textBeforeCursor As String
    
    currentPosition = Selection.Start
    
    lineStart = Selection.Paragraphs(1).Range.Start
    
    textBeforeCursor = Mid(Selection.Paragraphs(1).Range.Text, 1, currentPosition - lineStart)
    
    If Trim(textBeforeCursor) <> "" Then
        Exit Sub
    End If
    
    currentLine = Selection.Information(wdFirstCharacterLineNumber)
    
    If currentLine = 1 Then
        Exit Sub
    End If
    
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine
    
    prevLineEnd = Selection.Start

    Selection.Start = currentPosition
    Selection.End = prevLineEnd
    Selection.Delete
End Sub

Sub CutWithoutFinalLinekBreak()
    Dim sel As Range
    Set sel = Selection.Range
    
    If sel.End = sel.Paragraphs.Last.Range.End Then
        sel.End = sel.End - 1
    End If
    
    sel.Cut
End Sub
