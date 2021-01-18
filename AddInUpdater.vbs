Const GetAddInFrom = "\\192.168.x.xx\update\test.xlam"
Const myFile = "test.xlam"

Call DeactivateAddIn
Call UpdateAddIn
Call ActivateAddIn

MsgBox "アドインを更新しました"


Sub DeactivateAddIn
	Set Excel = GetObject(, "Excel.Application")
	For Each x In Excel.AddIns
		If x.Name = myFile Then
			x.Installed = False
		End If
	Next
End Sub

Sub ActivateAddIn
	Set Excel = GetObject(, "Excel.Application")
	For Each x In Excel.AddIns
		If x.Name = myFile Then
			x.Installed = True
		End If
	Next
End Sub

Sub UpdateAddIn
	Set WS = WScript.CreateObject("WScript.Shell")
	UserProfile = WS.ExpandEnvironmentStrings("%USERPROFILE%")
	Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
	FSO.CopyFile GetAddInFrom, UserProfile & "\AppData\Roaming\Microsoft\AddIns\", True
End Sub