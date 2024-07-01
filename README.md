# AllWebScraping_
All the program for take the data and create the Excel with the input of data in the excel

You have to install VisualStudio with the follow extensions :

- In desktop and mobile : Development in .NET desktop



The file is in Console_WebScrap

When you open the folder, there is 4 folder :
- "Class_Webscrap"
- "Console_WebScraping"
- "ExcelLocalBiblioC"
- "Json_Files"

You have to install PowerShell and do the following steps :

- Install PowerShell : https://github.com/PowerShell/PowerShell/releases/download/v7.4.3/PowerShell-7.4.3-win-x64.zip
- Run PowerShell
- Go to your file location of the solution with PowerShell, type "cd (and your file path)" to the Console_Webcrap folder
- Type in PowerShell the following line : "pwsh bin/Debug/netX/playwright.ps1 install"
- "netX" has to be change with your current version of .NET, you can see it when you click on bin and than Debug, the folder netX will appear.

You have to open Class_Webscrap and after launch Class_Webscrap.sln.

If the launch program isn't Console_Webscraping.cs, follow the next steps :

- Do a right click on the project / folder Console_WebScraping in Visual Studio (on the right of your screen in file explorer)
- Click on "Define at startup project"

Launch the program.

If you have errors, just try to wait a moment.

If the errors remains, try to look at the dependances :

- Richt click on the solution in Visual Studio (on the right of your screen in file explorer)
- Click on "Manage NuGet package"

Search for the following packages :

- IronXL.Excel
- Microsoft.Playwright
- Microsoft.Playwright.NUnit
- Selenium.WebDriver
- Selenium.WebDriver.ChromeDriver

when i'ts done you can finally run the program.

Wait.

When the command prompt say done at the end, you can open the excel file and close all the windows open.


If you want to change the path of the file or the default name, follow the nexts steps :

- Go to Program.cs of Console_WebScraping and change the "path" variable, you can do a ctrl + F for search it. 

This is for changing the path of the Excel file and the name of the file.

For exemple : E:\\Stage\\Virtual_game\\WebScraping\\dataFormatts.xlsx create the file dataFormatts the the CD E: and the folder WebScraping.
