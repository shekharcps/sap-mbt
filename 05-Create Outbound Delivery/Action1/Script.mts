Browser("Home").Page("Home").SAPUIButton("More groups").Click @@ script infofile_;_ZIP::ssf1.xml_;_
Browser("Home").Page("Home").SAPUIList("SAPUIList").Select "Outbound Delivery Creation" @@ script infofile_;_ZIP::ssf3.xml_;_
Browser("Home").Page("Home").SAPUITile("SAPUITile").Click @@ script infofile_;_ZIP::ssf4.xml_;_
Browser("Home").Page("Home").SAPUIButton("Expand Header").Click @@ script infofile_;_ZIP::ssf5.xml_;_
wait(5)
orderNumber = Parameter("orderNumber")
While Not Browser("Home").Page("Home").SAPUITable("Sales Orders Due for Delivery").RowCount > 1
	wait(1)
Wend
orderRowNum = Browser("Home").Page("Home").SAPUITable("Sales Orders Due for Delivery").FindRowByCellContent(2,orderNumber, 1)
Browser("Home").Page("Home").SAPUITable("Sales Orders Due for Delivery").SelectRow orderRowNum
Browser("Home").Page("Home").SAPUIButton("Create Deliveries").Click
property_DisplayLog = Browser("Home").Page("Home").SAPUIButton("Display Log").GetROProperty("disabled")
If property_DisplayLog = "False" Then
	Reporter.ReportEvent micPass, "Display Log on Create Delivery", "Display Log button gets enabled"
	Browser("Home").Page("Home").SAPUIButton("Display Log").Click
	deliveryNumber = Browser("Home").Page("Home").SAPUITable("Deliveries").GetCellData(1,"Delivery") 'Delivery
	Parameter("deliveryNumber") = deliveryNumber
	Reporter.ReportEvent micPass, "Order Delivery Number", "Order delivery number is : "&deliveryNumber
	Browser("Home").Page("Home").SAPUIButton("Close").Click
	Browser("Home").Page("Home").Image("Company Logo").Click
Else
	Reporter.ReportEvent micFail, "Display Log on Create Delivery", "Display Log button gets enabled"
End If

