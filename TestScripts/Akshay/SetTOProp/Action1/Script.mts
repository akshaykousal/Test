SystemUtil.Run "C:\Program Files\HP\QuickTest Professional\samples\flight\app\flight4a.exe"
Dialog("Login").Activate
Dialog("Login").WinEdit("Agent Name:").Set "Akshay"
Dialog("Login").WinEdit("Agent Name:").SetTOProperty "attached text","Password:"
Dialog("Login").WinEdit("Agent Name:").Set "Mercury"
Dialog("Login").WinButton("OK").Click