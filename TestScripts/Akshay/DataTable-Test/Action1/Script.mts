SystemUtil.Run "C:\Program Files\HP\QuickTest Professional\samples\flight\app\flight4a.exe"
Dialog("Login").Activate
Set oWinButton = Description.Create
oWinButton("micclass").Value = "WinButton"
'oWinButton("text").Value = "OK"

Dialog("Login").WinEdit("Class Name:=WinEdit", "attached text:=Agent Name:").Set DataTable.Value ("A",dtGlobalSheet)
Dialog("Login").WinEdit("Class Name:=WinEdit", "attached text:=Password:").Set DataTable.Value ("B",dtGlobalSheet)

Set AllWinButtons = Dialog("Login").ChildObjects(oWinButton)
For ii = 0 to AllWinButtons.Count-1
	If AllWinButtons(ii).GetROProperty("text") = "OK" Then
		AllWinButtons(ii).Click
		Exit For
	End If
Next
Set AllWinButtons = Nothing

WindowExist = Window("Flight Reservation").Exist(5)

If WindowExist Then
	Set AllWinButtons = Window("Flight Reservation").ChildObjects(oWinButton)
	For ii = 0 to AllWinButtons.Count-1
		If AllWinButtons(ii).GetROProperty("text") = "&Update Order" Then
			msgbox "Update Order button status -->" & AllWinButtons(ii).GetROProperty("enabled")
			Exit For
		End If
	Next
Else
	msgbox "Window(Flight Reservation) did not show up in 10 secs"
End If

