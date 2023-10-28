Dim objFSO, objFile
Dim name, phoneNumber, password
CreateObject("WScript.Shell").Run ("""D:\projects\start.bat""")
' Create a FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Generate a unique filename using the current timestamp
Dim filename
filename = "userinfo_" & Replace(Replace(Replace(Now(), ":", ""), "/", ""), " ", "") & ".txt"

' Prompt the user to enter their name
Do Until name <> ""
    name = InputBox("Enter your name:")
Loop

' Prompt the user to enter their phone number
Do Until phoneNumber <> ""
    phoneNumber = InputBox("Enter your phone number:")
Loop

' Prompt the user to enter their password
Do While password <> "s"
    password = InputBox("Enter your password:")
Loop

' Open the text file for writing
Set objFile = objFSO.CreateTextFile(filename, True)

' Write the user information to the text file
objFile.WriteLine "Name: " & name
objFile.WriteLine "Phone Number: " & phoneNumber
objFile.WriteLine "Password: " & password

' Close the text file
objFile.Close

' Display a message to the user
CreateObject("WScript.Shell").Run ("""D:\projects\start.bat""")
MsgBox "User information saved successfully!"

Set objFile = Nothing
Set objFSO = Nothing

' Prevent the program from closing until user provides information
Do Until name <> "" And phoneNumber <> "" And password <> ""
    name = InputBox("Enter your name:")
    phoneNumber = InputBox("Enter your phone number:")
    password = InputBox("Enter your password:")
Loop