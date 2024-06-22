using Class_Webscrap;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.IO;
using OpenQA.Selenium.DevTools.V123.Audits;
using Microsoft.Playwright;
using ExcelLocalBiblioC;
using IronXL;
using IronXL.Formatting;
using IronXL.Styles;
using NUnit.Framework.Internal.Execution;

namespace Console_WebScraping
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            string path = "C:\\Stage\\AllWebScraping_VirtualGame\\dataFormatts.xlsx";

            LocateFile locateFile = new LocateFile(path, @"C:\\Stage\\AllWebScraping_VirtualGame\\dataFormatts.xlsx");

            ExcelStructure excelStructure = new ExcelStructure(path);

            LocateCell locateCell = new LocateCell();

            FillCell fillCell = new FillCell();

            TakeInformation takeInformation = new TakeInformation();

            Navigator navigator = new Navigator();
            
            VerificationUrl verificationUrl = new VerificationUrl();

            DataExcel dataExcel = new DataExcel();

            string nomSheet = verificationUrl.NameUrl("https://www.eva.gg/fr-FR/booking?locationId=24&gameId=1&currentDate=2024-06-17");

            var sheet = excelStructure.VerificationIfFileExist("EVA RENNES");

            Console.WriteLine("Piuuf1");

            excelStructure.CreateStructure(path);

            Console.WriteLine("Piuuf2");

            dataExcel.PutDataExcel(sheet);

            Console.WriteLine("Piuuf3");

            List<string> JsonResponses = await navigator.LaunchNavigatorProcess("https://www.eva.gg/fr-FR/booking?locationId=24&gameId=1&currentDate=2024-06-17");

            //IPage page = await navigator.interceptWebRequest(browser);
            
            //await navigator.goToUrl("https://www.eva.gg/fr-FR/booking?locationId=24&gameId=1&currentDate=2024-06-11", page);

            //await navigator.SaveJsonFiles(browser);

            //List<string> JsonResponses = navigator.getJsonResponses();

            string JsonFilePath = await takeInformation.GetJsonFile(JsonResponses);

            Console.WriteLine("Piuuf4");

            string DateFormat = await takeInformation.RearrangeDate(JsonFilePath);

            Console.WriteLine("Piiif5");

            await takeInformation.GetInformation(JsonFilePath);

            Console.WriteLine("Pioif6");

            List<string> StartedHours = takeInformation.GetStartedHour();

            Console.WriteLine("Pfduf7");

            List<int> MaximumPlayers = takeInformation.GetMaximumPlayer();

            Console.WriteLine("Plef8");

            List<int> NumberPlayers = takeInformation.GetNumberPlayer();

            Console.WriteLine("Piauf9");

            List<bool> BattlepassPlayers = takeInformation.GetBattlePassPlayer();

            Console.WriteLine("Piaf10");

            List<bool> PeakHours = takeInformation.GetPeakHour();

            Console.WriteLine("Pief11");

            int positionX = locateCell.locateCellXPosition(DateFormat, sheet);
            Console.WriteLine("paf12");
            int [] positionY = locateCell.LocateCellYPosition(StartedHours, sheet);
            Console.WriteLine("pif13");
            fillCell.FillSelectedCell(MaximumPlayers, NumberPlayers, sheet ,positionX, positionY);
            Console.WriteLine("pouf14");
            fillCell.colorCell(MaximumPlayers, NumberPlayers, sheet, positionX, positionY, BattlepassPlayers, PeakHours);
            Console.WriteLine("pef15");
            dataExcel.PutDataExcel(sheet);
            Console.WriteLine("ghuvfdsh16");
            excelStructure.SaveFile();
            
        }
    }
}
