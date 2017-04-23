SystemUtil.Run "C:\Program Files\HP\QuickTest Professional\samples\flight\app\flight4a.exe","","C:\Program Files\HP\QuickTest Professional\samples\flight\app\",""
Dialog("Login").Activate
Dialog("Login").WinEdit("Agent Name:").Set "aks"
Dialog("Login").WinEdit("Password:").SetSecure "583efeca04606f1aff866aead76c4d9b440b85bf"
Dialog("Login").WinButton("OK").Click

If Window("Flight Reservation").Exist (15) Then
	Result = "Pass"
	Window("Flight Reservation").Close
Else
	Result = "Fail"
	Dialog("Login").WinButton("Cancel").Click
End If

Msgbox Result
