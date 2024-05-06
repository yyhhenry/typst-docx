```vba
Sub PasteTypstDocx()
    Dim http As Object
    Dim url As String
    Dim result As String
    Dim timestamp As String

    Shell "typst-docx.exe", vbMinimizedNoFocus
    
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
        Selection.InsertFile FileName:=result
    End If
    
    Set http = Nothing
End Sub
```
