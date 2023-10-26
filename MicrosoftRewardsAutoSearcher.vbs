Function GenerateRandomQuery()
    Dim characters, query, i, queryLength
    characters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-!@#$%^&*()+=><{}[];:"
    queryLength = Int((21 - 7 + 1) * Rnd + 7)
    
    For i = 1 To queryLength
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

Dim urls(34), i, url
Randomize
For i = 0 To 34
    urls(i) = GenerateRandomURL()
Next

Set shell = CreateObject("WScript.Shell")
For Each url In urls
    shell.Run "cmd /c start " & url
    Dim sleepDuration
    sleepDuration = Int((9000 - 5000 + 1) * Rnd + 5000)
    WScript.Sleep sleepDuration
Next
