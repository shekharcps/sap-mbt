orderNumber = Parameter("orderNumber")
With Browser("Home")
	.Page("Home").WebElement("tabSalesOrders").Click
	.Page("Home").WebElement("manageSalesOrders").Click
	wait(10)
	While Not .Page("Home").SAPUITable("Sales Orders").RowCount > 1
		wait(1)
	Wend
	wait(10)
	orderRowNum = .Page("Home").SAPUITable("Sales Orders").FindRowByCellContent(1,orderNumber, 1)
	.Page("Home").SAPUITable("Sales Orders").SelectRow orderRowNum
	.Page("Home").SAPUITable("Sales Orders").SelectItemInCell orderRowNum,"Sales Order", orderNumber @@ script infofile_;_ZIP::ssf2.xml_;_
	.Page("Home").Link("Display Sales Order -").Click ' Display Sales Order VA03 @@ script infofile_;_ZIP::ssf5.xml_;_
	.Page("Home").SAPFrame("Display Standard Order").SAPButton("Header Output Preview").Click @@ script infofile_;_ZIP::ssf7.xml_;_
	.Page("Home").SAPFrame("Display Standard Order").SAPButton("Display Document").Click
	If .Page("Display Standard Order").SAPFrame("Display Standard Order").WebTable("Display Messages").Exist(2) Then
		Reporter.ReportEvent micWarning, "Display Messages", "Displayed Error Messages"
		.Page("Display Standard Order").SAPFrame("Display Standard Order").SAPButton("Close").Click
		.Page("Display Standard Order").Image("Company Logo").Click
	End If
End With
