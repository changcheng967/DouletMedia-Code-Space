Set objShell = CreateObject("WScript.Shell")
tries = 10

' Display the initial message box with options
do
    response = MsgBox("Give money or Exit?", vbYesNo + vbQuestion, "Support Me")
    
    If response = vbYes Then
        ' Redirect to Buy Me a Coffee page
        objShell.Run "https://buymeacoffee.com/changcheng967"
        ' Continue the loop so the window doesn't exit after clicking "Give money"
    ElseIf response = vbNo Then
        ' Start the password prompt if Exit is chosen
        Do
            password = InputBox("Enter password to exit (" & tries & " tries left):", "Password Required")
            
            If password = "frankbalabala" Then
                MsgBox "Correct password. Exiting...", vbInformation, "Access Granted"
                WScript.Quit
            Else
                tries = tries - 1
                MsgBox "Wrong password. " & tries & " tries left.", vbExclamation, "Access Denied"
            End If
            
            ' If the user runs out of attempts, shut down the computer
            If tries = 0 Then
                MsgBox "Too many incorrect attempts. Shutting down...", vbCritical, "Locked Out"
                objShell.Run "shutdown /s /t 0", 0, False
                WScript.Quit
            End If
        Loop While tries > 0
    End If
Loop
