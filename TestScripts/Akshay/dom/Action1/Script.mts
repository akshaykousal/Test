'http://www.espncricinfo.com/rankings/content/page/211271.html

Set oPage = Browser("ICC rankings - Test, ODI,").Page("ICC rankings - Test, ODI,")
Set oDetailsIndiaRow = oPage.Object.getElementsByClassName("pnl650M")
Set oDetailsIndiaRowGrid = oDetailsIndiaRow.Object.getElementsByClassName("StoryengineTable")

Set oPage = Nothing
Set oDetailsIndiaRow = Nothing
Set oDetailsIndiaRowGrid = Nothing


Browser("ICC rankings - Test, ODI,").Page("ICC rankings - Test, ODI,").Object.RunScript("alert('Hello There!!');")


'getElementsByClassName("pnl650M")



