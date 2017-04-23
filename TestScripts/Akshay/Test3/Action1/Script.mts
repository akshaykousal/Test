SystemUtil.Run "C:\Program Files\Internet Explorer\iexplore.exe", "https://in.yahoo.com/"
Set dBrowser = Description.Create
dBrowser("micclass").value = "Browser"
dBrowser("name").value = "Yahoo"

Set dPage = Description.Create
dPage("micclass").value = "Page"
dPage("title").value = "Yahoo"

Set dLink = Description.Create
dLink("micclass").value = "Link"
'dLink("text").value = "Raj Babbar UP chief"
myLinks = ""
Browser(dBrowser).Page(dPage).Sync
Set oLinks = Browser(dBrowser).Page(dPage).ChildObjects(dLink)
For i=0 to oLinks.Count-1
myLinks =myLinks &  i & "-->" & oLinks(i).GetROProperty("text") &vbNewline
'	If  oLinks(i).GetROProperty("text") = "Cricket" Then
'		msgbox "Cricket link found. Clicking the link"
'		oLinks(i).Click
'		Exit For
'	End If
Next
print myLinks
