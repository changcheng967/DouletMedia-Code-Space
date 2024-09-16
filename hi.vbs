Dim attempts, maxAttempts, password, triesLeft

attempts = 0
maxAttempts = 10

' Create popup for asking to give money
answer = MsgBox("Give money", vbYesNo + vbInformation, "Prompt")

If answer = vbYes Then
    ' Redirect to Buy Me a Coffee link but don't exit the script
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.Run "https://buymeacoffee.com/changcheng967"
End If

' Continue to password prompt even after clicking "Give money"
Do
    triesLeft = maxAttempts - attempts
    password = InputBox("Enter password (" & triesLeft & " attempts left):", "Password Prompt")
    
    If password = "frankbalabala" Then
        MsgBox "Exiting..."
        WScript.Quit
    Else
        attempts = attempts + 1
        If attempts >= maxAttempts Then
            MsgBox "Wrong password entered 10 times. Shutting down..."
            Set WshShell = CreateObject("WScript.Shell")
            WshShell.Run "shutdown /s /t 0"
            WScript.Quit
        Else
            MsgBox "Wrong password. " & (maxAttempts - attempts) & " attempts left."
        End If
    End If
Loop
