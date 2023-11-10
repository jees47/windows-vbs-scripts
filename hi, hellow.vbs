Dim messages(3)
messages(0) = "hi"
messages(1) = "hey"
messages(2) = "heyo"
messages(3) = "hello"
currentMessageIndex = 0

Do
    MsgBox messages(currentMessageIndex), vbInformation, "Popup Message"
    
    ' Cycle through the messages in the array
    currentMessageIndex = (currentMessageIndex + 1) Mod 4
Loop