using Class_Webscrap;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.IO;
using OpenQA.Selenium.DevTools.V123.Audits;
using Microsoft.Playwright;
using ExcelLocalBiblioC;
using IronXL;

namespace Console_WebScraping
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            string path = "N:\\Stage\\Virtual_game\\Webcsrap\\dataFormatts.xlsx";

            LocateFile locateFile = new LocateFile(path, @"N:\Stage\Virtual_game\Webcsrap\AllWebScraping_VirtualGame\dataFormatts.xlsx");

            ExcelStructure excelStructure = new ExcelStructure(path);

            LocateCell locateCell = new LocateCell();

            FillCell fillCell = new FillCell();

            TakeInformation takeInformation = new TakeInformation();

            Navigator navigator = new Navigator();
            
            VerificationUrl verificationUrl = new VerificationUrl();

            DataExcel dataExcel = new DataExcel();

            string nomSheet = verificationUrl.NameUrl("https://www.eva.gg/fr-FR/booking?locationId=24&gameId=1&currentDate=2024-06-17");

            var sheet = excelStructure.VerificationIfFileExist("EVA RENNES");

            Console.WriteLine("Piuuf");

            excelStructure.CreateStructure(path);

            Console.WriteLine("Piuuf");

            dataExcel.PutDataExcel(sheet);

            Console.WriteLine("Piuuf");

            List<string> JsonResponses = await navigator.LaunchNavigatorProcess("https://www.eva.gg/fr-FR/booking?locationId=24&gameId=1&currentDate=2024-06-17");

            //IPage page = await navigator.interceptWebRequest(browser);
            
            //await navigator.goToUrl("https://www.eva.gg/fr-FR/booking?locationId=24&gameId=1&currentDate=2024-06-11", page);

            //await navigator.SaveJsonFiles(browser);

            //List<string> JsonResponses = navigator.getJsonResponses();

            string JsonFilePath = await takeInformation.GetJsonFile(JsonResponses);

            Console.WriteLine("Piuuf");

            string DateFormat = await takeInformation.RearrangeDate(JsonFilePath);

            Console.WriteLine("Piiif");

            await takeInformation.GetInformation(JsonFilePath);

            Console.WriteLine("Pioif");

            List<string> StartedHours = takeInformation.GetStartedHour();

            Console.WriteLine("Pfduf");

            List<int> MaximumPlayers = takeInformation.GetMaximumPlayer();

            Console.WriteLine("Plef");

            List<int> NumberPlayers = takeInformation.GetNumberPlayer();

            Console.WriteLine("Piauf");

            List<bool> BattlepassPlayers = takeInformation.GetBattlePassPlayer();

            Console.WriteLine("Piaf");

            List<bool> PeakHours = takeInformation.GetPeakHour();

            Console.WriteLine("Pief");

            int positionX = locateCell.locateCellXPosition(DateFormat, sheet);
            Console.WriteLine("paf");
            int [] positionY = locateCell.LocateCellYPosition(StartedHours, sheet);
            Console.WriteLine("pif");
            fillCell.FillSelectedCell(MaximumPlayers, NumberPlayers, sheet ,positionX, positionY);
            Console.WriteLine("pouf");
            fillCell.colorCell(MaximumPlayers, NumberPlayers, sheet, positionX, positionY, BattlepassPlayers, PeakHours);
            Console.WriteLine("pef");
            excelStructure.SaveFile(path);
            
        }
    }
}
