wait(5)
Browser("Home").Page("Home").SAPUIButton("More groups").Highlight @@ script infofile_;_ZIP::ssf13.xml_;_
Browser("Home").Page("Home").SAPUIButton("More groups").HoverTap
Browser("Home").Page("Home").SAPUIButton("More groups").Click @@ script infofile_;_ZIP::ssf14.xml_;_
 @@ script infofile_;_ZIP::ssf15.xml_;_
'While Not Browser("Home").Page("Home").SAPUIButton("More groups").Exist(1)  
'	wait(1)
'	If Browser("Home").Page("Home").SAPUIButton("More groups").Exist(1) Then
'		propertyMoreGroups = Browser("Home").Page("Home").SAPUIButton("More groups").GetROProperty("disabled")
'		If propertyMoreGroups = "False" Then
'			Browser("Home").Page("Home").SAPUIButton("More groups").Click
'		Else
'			wait(5)
'		End If
'	End If
'Wend

Browser("Home").Page("Home").WebElement("Stock Monitoring").Click
Browser("Home").Page("Home").Link("StockMultiple Materials").Click
Browser("Home").Page("Home").SAPUIButton("Expand Header").Click
wait(5)
Browser("Home").Page("Home").WebEdit("Material Number").Set "EWMS4-01"
pressButton("TAB")
wait 5
Browser("Home").Page("Home").WebEdit("Material Number").Set "EWMS4-02"
pressButton("TAB")
'Browser("Home").Page("Home").SAPUITextEdit("Storage Location").Set "171S" @@ script infofile_;_ZIP::ssf4.xml_;_
Browser("Home").Page("Home").WebEdit("Storage Location").Set "171S"
pressButton("TAB")
Browser("Home").Page("Home").SAPUIButton("Go").Click @@ script infofile_;_ZIP::ssf11.xml_;_
wait(5)

'	Check Stock Availability
For i = 2 To 3
	materialNumber = Browser("Home").Page("Home").WebTable("tblMaterials").GetCellData(i,1)
	blockedStock = Browser("Home").Page("Home").WebTable("tblMaterials").GetCellData(i,11)
	arrblockedStock = split(blockedStock, ".")
	blockedStock = arrblockedStock(0)
	If blockedStock > 1 Then
		Reporter.ReportEvent micPass, "Blocked Stock is greater than 1", "Blocked Stock of material "&materialNumber& " is: "&blockedStock
		If materialNumber = "EWMS4-02" Then
			blockedStock_EWMS4_02 = blockedStock
			Parameter("blockedStock_EWMS4_02") = blockedStock_EWMS4_02
		ElseIf materialNumber = "EWMS4-01" Then 
			blockedStock_EWMS4_01 = blockedStock
			Parameter("blockedStock_EWMS4_01") = blockedStock_EWMS4_01
		End If
	Else
		Reporter.ReportEvent micFail, "Stock is not available", "Blocked Stock of material "&materialNumber& " is: "&blockedStock
		ExitAction
	End If
Next





'Function - Press Button {ENTER} / {TAB}
Public Sub pressButton(buttonName)
	Set WinShell = CreateObject("WScript.Shell")
	MyVar = Ucase (buttonName)
   	Select Case MyVar
      		Case "ENTER"
      			WinShell.SendKeys "{ENTER}"   	
		Case "TAB"
      			WinShell.SendKeys "{TAB}"        			
       	Case Else
   	End Select
   	Set WinShell = Nothing
End Sub
