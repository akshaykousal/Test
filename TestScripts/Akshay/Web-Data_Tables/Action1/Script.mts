'*********WebTable************
SystemUtil.Run "C:\Program Files\Internet Explorer\iexplore.exe", "C:\Users\Akshay\Desktop\WebTable.html"
Browser("Browser").Page("Page").Sync
Browser("Browser").Page("Page").WebTable("Google.com")

Set oWebtable = Description.Create
oWebTable("micclass").Value="WebTable"
oWebTable("name").Value="Google.com"

Set oAllWebTables = Browser("Browser").Page("Page").ChildObjects(oWebTable)

Set MyWebTable = oAllWebTables(0)
Set MyWebTable = Browser("Browser").Page("Page").WebTable("Google.com")
RowCount = MyWebTable.RowCount
For rowno = 1 to RowCount
	ColCount = MyWebTable.ColumnCount(rowno)
	For colno = 1 to ColCount
		RowColCellData = MyWebTable.GetCellData(rowno,colno)
		LinkCount = MyWebTable.ChildItemCount(rowno,colno,"Link")
			If LinkCount > 0 Then
				set LinkInCell = MyWebTable.ChildItem(rowno,colno,"Link",LinkCount-1)
					If LinkInCell.GetROProperty("text") = "Facebook.com" Then
						LinkInCell.Click
					End If
			End If
	Next
Next


''*********DataTable************

'msgbox DataTable.GetSheet("Global").GetParameter("B").ValueByRow(1)
DataTable.AddSheet("MySheet")
DataTable.GetSheet("MySheet").AddParameter "Col3","Test"

RowCnt = DataTable.GetSheet("Global").GetRowCount
ColCnt = DataTable.GetSheet("Global").GetParameterCount

For rowno = 1 to RowCnt
	DataTable.GetSheet("Global").SetCurrentRow(rowno)
	For colno = 1 to ColCnt
		msgbox "(" & rowno & "," & colno &") --> " & DataTable.Value(colno,"Global")
	Next
Next


msgbox DataTable.GetSheet("Global").GetParameter("Col2").ValueByRow(3)


msgbox ColAvailable("Col3")

Function ColAvailable(ColName)
	On Error Resume Next
	ColAvailable = True
	Err.Clear
	DataTable.GetSheet("Global").GetParameter(ColName)
	If Err.Number <>0 Then
		ColAvailable = False
	End If
End Function

DataTable.GetSheet("Global").SetCurrentRow(5)
RV = DataTable.RawValue("Col2","Global")
V = DataTable.Value("Col2","Global")
Msgbox "RV --> "& RV & ". V --> " & V