Set WshShell = WScript.CreateObject("WScript.Shell")

    WScript.Sleep 3000

For i = 1 To 5
    Dim randomString
    randomString = GenerateRandomString(5)

    WshShell.SendKeys randomString
    WshShell.SendKeys "{ENTER}"

    WScript.Sleep 2000
Next

Function GenerateRandomString(length)
    Dim characters, i
    characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    randomString = ""
    For i = 1 To length
        randomString = randomString & Mid(characters, Int((Len(characters) * Rnd) + 1), 1)
    Next
    GenerateRandomString = randomString
End Function
