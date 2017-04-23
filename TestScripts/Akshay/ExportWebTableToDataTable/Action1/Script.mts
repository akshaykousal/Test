SystemUtil.Run "C:\Program Files\Internet Explorer\iexplore.exe","C:\Users\Akshay\Desktop\WebTable.html"
Set MyWebTable = Browser("Browser").Page("Page").WebTable("Google.com")
HeaderRow = True
WebTableRowCount = MyWebTable.RowCount
For WebTableRowNo = 1 to WebTableRowCount
	DataTable.SetCurrentRow(WebTableRowNo)
	WebTableColCount = MyWebTable.ColumnCount(WebTableRowNo)
	For WebTableColNo = 1 to WebTableColCount
		If HeaderRow Then
			DataTable.GetSheet("Global").AddParameter "Col" & WebTableColNo,""
		End If
		DataTable.GetSheet("Global").GetParameter(WebTableColNo).value = MyWebTable.GetCellData(WebTableRowNo, WebTableColNo)
	Next
	HeaderRow = False
Next





