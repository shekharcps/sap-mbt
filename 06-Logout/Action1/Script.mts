With Browser("Home").Page("Home")
	.Image("Navigate to Home Page").Click
	.WebElement("Profile").Click
	.WebElement("Sign Out").Click
	.WebElement("msg_Ok_Btn").Click
End With

