Set WshShell = WScript.CreateObject("WScript.Shell")

' Repeat the process 50 times
For i = 1 To 50
    ' Generate a random 5-character string
    Dim randomString
    randomString = GenerateRandomString(5)

    ' Send the random string to the Start Menu search box
    WshShell.SendKeys "^{ESC}" ' Open the Start Menu
    WScript.Sleep 1000 ' Sleep for a moment to ensure the Start Menu is open
    WshShell.SendKeys randomString ' Type the random string
    WScript.Sleep 1000 ' Sleep for a moment to display search results
    WshShell.SendKeys "{ENTER}" ' Press the Enter key

    ' Wait for some time before the next iteration
    WScript.Sleep 2000

    ' Close the Start Menu
    WshShell.SendKeys "{ESC}"
    WScript.Sleep 1000
Next

' Function to generate a random string
Function GenerateRandomString(length)
    Dim characters, i
    characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    randomString = ""
    For i = 1 To length
        randomString = randomString & Mid(characters, Int((Len(characters) * Rnd) + 1), 1)
    Next
    GenerateRandomString = randomString
End Function
