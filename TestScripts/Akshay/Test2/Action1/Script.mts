'DataTable.AddSheet "Login"
'DataTable.ImportSheet "C:\Users\Akshay\Desktop\myExcel.xls",1,"Login"
'RowCnt = DataTable.GetSheet("Login").GetRowCount
'
'For ii = 1 to RowCnt
'DataTable.SetCurrentRow(ii)
''DataTable.GetSheet "Login"
'SystemUtil.Run "C:\Program Files\HP\QuickTest Professional\samples\flight\app\flight4a.exe","","C:\Program Files\HP\QuickTest Professional\samples\flight\app\",""
'Dialog("Login").Activate
'Dialog("Login").WinEdit("Agent Name:").Set DataTable(1,"Login")
'Dialog("Login").WinEdit("Password:").Set DataTable(2,"Login")
'Dialog("Login").WinButton("OK").Click
'If Window("Flight Reservation").Exist (12)	 Then
'Window("Flight Reservation").Close
'DataTable(3,3) = "Pass"
'Else
'SystemUtil.CloseDescendentProcesses 'Closes all processes opened by QTP/UFT
'DataTable(3,3) = "Fail"
'End If
'
'Next
'
'
'
'
'
'
'
msgbox "b"
