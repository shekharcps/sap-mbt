orderNumber = Parameter("orderNumber")
Browser("Home").Page("Home").WebElement("tabSalesOrders").Click
Browser("Home").Page("Home").WebElement("manageSalesOrders").Click
orderRowNum = Browser("Home").Page("Home").SAPUITable("Sales Orders").FindRowByCellContent(1,orderNumber, 1)
Browser("Home").Page("Home").SAPUITable("Sales Orders").SelectRow orderRowNum
Browser("Home").Page("Home").SAPUITable("Sales Orders").SelectItemInCell orderRowNum,"Sales Order", orderNumber @@ script infofile_;_ZIP::ssf2.xml_;_
Browser("Home").Page("Home").Link("Display Sales Order -").Click ' Display Sales Order VA03 @@ script infofile_;_ZIP::ssf5.xml_;_
Browser("Home").Page("Home").SAPFrame("Display Standard Order").SAPButton("Header Output Preview").Click @@ script infofile_;_ZIP::ssf7.xml_;_
Browser("Home").Page("Home").SAPFrame("Display Standard Order").SAPButton("Display Document").Click
If Browser("Home").Page("Display Standard Order").SAPFrame("Display Standard Order").WebTable("Display Messages").Exist(2) Then
	Reporter.ReportEvent micWarning, "Display Messages", "Displayed Error Messages"
	Browser("Home").Page("Display Standard Order").SAPFrame("Display Standard Order").SAPButton("Close").Click
	Browser("Home").Page("Display Standard Order").Image("Company Logo").Click
End If
