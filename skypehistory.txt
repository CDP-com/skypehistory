'VBScript: skypehistory.vbs
' Define global variables
Dim oFSO, chat_file, folder_to_save
' Directory where You want to save history (you can modify it)
' Now it is relative, so it will be created where Your *.vbs script runs
' Added path to desktop - Marshall Stiles
folder_to_save = "%USERPROFILE%\Desktop\SkypeChatHistory"
Set wshShell = CreateObject( "WScript.Shell" )
folder_to_save = wshShell.ExpandEnvironmentStrings(folder_to_save)
line_count = 0

On Error Resume Next
' Connect to Skype API via COM
Set oSkype = WScript.CreateObject("Skype4COM.Skype", "Skype_")
If Not IsObject(oSkype) Then
    WScript.Echo "Error: Unable to connect to Skype"
	WScript.Quit
End If

' Open skype, if it is not running
If Not oSkype.Client.IsRunning Then
	oSkype.Client.Start()

	Do While Not oSkype.Client.IsRunning
	  'nothing happening
	Loop

	Dim PauseTime, Start

	PauseTime = 20 'Set Duration
	Start = Timer 'Set start time

	Do While Timer < Start + PauseTime
	  'nothing happening
	Loop
	
End If


' Create FSO
Set oFSO = CreateObject("Scripting.FileSystemObject")
set_next_free_dir()

WScript.Echo "Skype history will be saved. Found " & oSkype.Chats.Count & " chat groups."

' Iterate chats
For Each oChat In oSkype.Chats
names = ""
' First name is You, so it is unnecessary to keep
no_1st_flag = TRUE
For Each oUser In oChat.Members
If no_1st_flag Then
no_1st_flag = FALSE
Else
   names = names & "_" & oUser.FullName
End If
Next
get_file("chat" & names & ".txt")
chat_file.WriteLine(vbNewLine & "==== CHAT HISTORY (" & Replace(names, "_", "") & ") ====" & vbNewLine)
line_count = line_count + oChat.Messages.Count
' Fix by an anonymous commenter
If oChat.Messages.Count > 0 Then
For Each oMsg In oChat.Messages
' Fix by Vadim Kravchenko
On Error Resume Next
chat_file.WriteLine(oMsg.FromDisplayName & " (" & oMsg.Timestamp & "): " & oMsg.Body)
Next
End If
chat_file.Close
Next

WScript.Echo "Backup was finished (" & line_count & " lines saved). You can find your chats in: " & folder_to_save

' Garbage collection
SET chat_file = NOTHING
SET folder_to_save = NOTHING
SET oFSO = NOTHING
SET oSkype = NOTHING
SET wshShell = NOTHING

' Access to a file given by name
Sub get_file(file_name)
' Parameter fix by: rommeech
Set chat_file = oFSO.OpenTextFile(folder_to_save & "/" & file_name, 8, True, -1)
End Sub

' Find an appropriate directory the logs to save, however, to avoid collision with former dirs
Sub set_next_free_dir()
If oFSO.FolderExists(folder_to_save) Then
ext = 1
While oFSO.FolderExists(folder_to_save & "_" & ext) And ext < 100
  ext = ext + 1
Wend
folder_to_save = folder_to_save & "_" & ext
End If
oFSO.CreateFolder(folder_to_save)
End Sub
