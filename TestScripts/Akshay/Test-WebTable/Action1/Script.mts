SystemUtil.Run "C:\Program Files\Internet Explorer\iexplore.exe", "C:\Users\Akshay\Desktop\WebTable.html"
Browser("Browser").Page("Page").Sync

Set oWebtable = Description.Create
oWebTable("micclass").Value="WebTable"
oWebTable("name").Value="Google.com"

Set oAllWebTables = Browser("Browser").Page("Page").ChildObjects(oWebTable)

Set MyWebTable = oAllWebTables(0)
blnRowColFound=False
RowCount = MyWebTable.RowCount
For rowno = 1 to RowCount
	ColCount = MyWebTable.ColumnCount(rowno)
	For colno = 1 to ColCount
		RowColCellData = MyWebTable.GetCellData(rowno,colno)
		LinkCount = MyWebTable.ChildItemCount(rowno,colno,"Link")
		If LinkCount > 0 Then
			For intIndex = 0 To LinkCount-1
				set LinkInCell = MyWebTable.ChildItem(rowno,colno,"Link",intIndex)
				If LinkInCell.GetROProperty("text") = "Facebook.com" Then
					LinkInCell.Click
					blnRowColFound = True
					Exit For
				End If
			Next
		End If
		If blnRowColFound Then
			Exit For
		End If
	Next
	If blnRowColFound Then
		Exit For
	End If
Next
