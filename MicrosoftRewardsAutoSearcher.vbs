Function GenerateRandomQuery()
    Dim characters, query, i
    characters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-"
    For i = 1 To 10
        query = query & Mid(characters, Int(Rnd() * Len(characters) + 1), 1)
    Next
    GenerateRandomQuery = query
End Function
Function GenerateRandomURL()
    Dim url, query
    query = GenerateRandomQuery()
    url = "https://www.bing.com/search?q=" & query
    GenerateRandomURL = url
End Function
Dim urls(29), i, url
Randomize
For i = 0 To 29
    urls(i) = GenerateRandomURL()
Next
Set shell = CreateObject("WScript.Shell")
For Each url In urls
    shell.Run "cmd /c start " & url
    WScript.Sleep 5000
Next
