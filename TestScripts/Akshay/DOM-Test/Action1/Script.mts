Call Fn_Stats("V Kohli","Ave")
Function Fn_Stats(strPlayer,strStatsOf)
	'SystemUtil.Run "http://stats.espncricinfo.com/ci/engine/records/averages/batting.html?class=1;id=2017;type=year"
	Set oPage = Browser("Cricket Records | Records").Page("Cricket Records | Records").WebElement("(function(d, s, id) {")
	Set oPageRecords = oPage.Object.all.tags("td")
	blnRecordFound = False

    Select Case strStatsOf
		Case "Mat"
			intIncrement = 1
		Case "Inns"
			intIncrement = 2
		Case "NO"
			intIncrement = 3
		Case "Runs"
			intIncrement = 4
		Case "HS"
			intIncrement = 5
		Case "Ave"
			intIncrement = 6
		Case "BF"
			intIncrement = 7
		Case "SR"
			intIncrement = 8
		Case "100"
			intIncrement = 9
		Case "50"
			intIncrement = 10
		Case "0"
			intIncrement = 11
		Case "4s"
			intIncrement = 12
		Case "6s"
			intIncrement = 13
	End Select
	
	For i = 0 to oPageRecords.length-1
		set CurrentRecord = oPageRecords(i)
		str = CurrentRecord.innerText
		If instr(str,strPlayer )>0 Then
			set CurrentRecord = Nothing
			set CurrentRecord = oPageRecords(i+intIncrement)
			str1 = CurrentRecord.innerText
			strRecordDetails = str &"--> " & strStatsOf & ": " & str1
			blnRecordFound = True
			Exit For
		End If
	Next

	If blnRecordFound Then
		msgbox strRecordDetails
	Else
		msgbox "No Record found by Name - " & strPlayer
	End If
	
	Set oPageRecords = Nothing
	Set oPage = Nothing
	set CurrentRecord = Nothing
End Function
