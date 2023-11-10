Dim age

' Prompt the user for their age
age = InputBox("Please enter your age:")

' Check if the user provided an age
If age <> "" Then
    ' Display the user's age
    MsgBox "Your age is: " & age
Else
    ' Display a message if the user did not enter an age
    MsgBox "You did not enter your age."
End If