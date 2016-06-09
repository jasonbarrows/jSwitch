'~ Toggle SPECIFIED NICs on or off
Option Explicit

Const NETWORK_CONNECTIONS = &H31&

Dim objShell, objFolder, objFolderItem, objEnable, objDisable
Dim folder_Object, target_NIC
Dim i, str_NIC, NIC, clsVerb
Dim str_NIC_Names, strEnable, strDisable
Dim bEnabled, bDisabled
Dim selected_NIC, current_path, filename, shortcut

str_NIC_Names = Array("Local Area Connection", "Wireless Network Connection")

strEnable = "En&able"
strDisable = "Disa&ble"

' Create objects and get items
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(NETWORK_CONNECTIONS)
Set objFolderItem = objFolder.Self
Set folder_Object = objFolderItem.GetFolder

' See if the namespace exists
If folder_Object Is Nothing Then
	Wscript.Echo "Could not find Network Connections"
	WScript.Quit
End If

For i = LBound(str_NIC_Names) to UBound(str_NIC_Names)
	Set target_NIC = Nothing

	' Look at each NIC and match to the chosen name
	For Each NIC In folder_Object.Items
		If LCase(NIC.Name) = LCase(str_NIC_Names(i)) Then
			Set target_NIC = NIC
		End If
	Next

	If target_NIC Is Nothing Then
		WScript.Echo "Unable to locate proper NIC"
		WScript.Quit
	End If

	bEnabled = True
	Set objEnable = Nothing
	Set objDisable = Nothing

	For Each clsVerb In target_NIC.Verbs
		If clsVerb.Name = strEnable Then
			Set objEnable = clsVerb
			bEnabled = False
		End If
		If clsVerb.Name = strDisable Then
			Set objDisable = clsVerb
		End If
	Next

	If bEnabled Then
		objDisable.DoIt
	Else
		objEnable.DoIt
		selected_NIC = str_NIC_Names(i)
	End If

	'~ Give the connection time to stop/start
	WScript.Sleep 1000
Next

' Set up environment for Shortcut
Set objShell = CreateObject("WScript.Shell")
current_path = objShell.CurrentDirectory
filename = objShell.SpecialFolders("Desktop") & "\jSwitch.lnk"
Set shortcut = objShell.CreateShortcut(filename)

If selected_NIC = "Local Area Connection" Then
	shortcut.IconLocation = current_path & "\wired.ico"
else
	shortcut.IconLocation = current_path & "\wireless.ico"
end if

shortcut.TargetPath = current_path & "\jSwitch.vbs"
shortcut.WorkingDirectory = current_path
shortcut.Save

WScript.Quit