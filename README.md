# excel post data to Google Spreadsheet
 Microsoft Excel post data to Google Spreadsheet with Selenium

![alt text](https://github.com/jenizar/excel-post-data-to-google-spreadsheet/blob/master/screenshot1.PNG)


![alt text](https://github.com/jenizar/excel-post-data-to-google-spreadsheet/blob/master/screenshot2.PNG)

Requirements:

Install selenium basic (see references)

step by step:
1. Create form and assign to spreadsheet name
2. Find variable name on each input text in form
3. Create html page in your localhost
3. Create data in excel workbook
4. Create button and assign macro each button 
5. Assign reference to Selenium Library and Insert macro coding

Sub Button1_Click()

Dim MyParm As New Webdriver

MyParm.Start "chrome", "http://localhost/getData/salesorder.html"
MyParm.get "http://localhost/getData/salesorder.html"
MyParm.Wait 500

MyParm.FindElementByName("entry.xxx").SendKeys (Range("a2").Value)
MyParm.FindElementByName("entry.xxx").SendKeys (Range("b2").Value)
MyParm.FindElementByName("entry.xxx").SendKeys (Range("c2").Value)
MyParm.FindElementByName("entry.xxx").SendKeys (Range("d2").Value)
MyParm.FindElementByName("entry.xxx").SendKeys (Range("e2").Value)
MyParm.FindElementByXPath("//button[@type='submit']").Click

End Sub

6. Test

References:
1. http://drvba.blogspot.com/
2. https://www.youtube.com/watch?v=2jZGhKugK70
