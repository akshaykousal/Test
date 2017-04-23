SystemUtil.Run "C:\Program Files\Internet Explorer\iexplore.exe","C:\Users\Akshay\Desktop\WebTable.html"
Set MyWebTable = Browser("Browser").Page("Page").WebTable("Google.com")

WebTableRowCount = MyWebTable.RowCount
For WebTableRowNo = 1 to WebTableRowCount
	WebTableColCount = MyWebTable.ColumnCount(WebTableRowNo)
	For WebTableColNo = 1 to WebTableColCount
		AllWebEditInCellCount = MyWebTable.ChildItemCount(WebTableRowNo,WebTableColNo,"WebEdit")
		If AllWebEditInCellCount >0 Then
			For AllWebEditInCellNo = 1 to AllWebEditInCellCount
				Set AllWebEditInCell = MyWebTable.ChildItem(WebTableRowNo,WebTableColNo,"WebEdit",AllWebEditInCellNo-1)
				AllWebEditInCell.Set "Text" & AllWebEditInCellNo
			Next
		End If
	Next
Next

a = Environment.Value("OS")
msgbox a
a = Environment("OS")
msgbox a
'a=1

SystemUtil.Run "C:\Program Files\Internet Explorer\iexplore.exe","C:\Users\Akshay\Desktop\WebTable.html"
Set MyWebTable = Browser("Browser").Page("Page").WebTable("Google.com")
Set oWebList = MyWebTable.ChildItem(3,2,"WebList",0)
oWebList.Select "Audi"
b = oWebList.GetRoProperty("value")
a=1





