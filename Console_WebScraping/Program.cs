using Class_Webscrap;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.IO;
using OpenQA.Selenium.DevTools.V123.Audits;
using Microsoft.Playwright;
using ExcelLocalBiblioC;
using IronXL;
using System.Text.Json;
using static System.Net.WebRequestMethods;
using System.Security.Policy;

namespace Console_WebScraping
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            UrlList urlList = new UrlList();

            urlList.MakeUrlList();

            List<string> Urllist = urlList.GetUrlList();
            string path = "..\\..\\..\\..\\..\\dataFormatts.xlsx";
            LocateFile locateFile = new LocateFile(path);
            ExcelStructure excelStructure = new ExcelStructure(path);
            LocateCell locateCell = new LocateCell();
            VerificationUrl verificationUrl = new VerificationUrl();
            NavigatorEVA navigatorEVA = new NavigatorEVA();
            NavigatorZL navigatorZL = new NavigatorZL();
            NavigatorOxmozvr navigatorOxmozvr = new NavigatorOxmozvr();
            

            TakeInformationEVA takeInformationEVA = new TakeInformationEVA();
            TakeInformationZL takeInformationZL = new TakeInformationZL();
            TakeInformationOxmozvr takeInformationOxmozvr = new TakeInformationOxmozvr();

            DataExcel dataExcel = new DataExcel();
            JsonFileList jsonFileList = new JsonFileList();

            string JsonFilePath = "";

            string date = DateTime.Now.ToString("dd/MM/yyyy");
            foreach (string url in Urllist)
            {
                string nomSheet = verificationUrl.NameUrl(url);

                if (nomSheet.Contains("EVA"))
                {
                    List<string> JsonResponses = await navigatorEVA.LaunchNavigatorProcess(url, nomSheet);

                    JsonFilePath = await takeInformationEVA.GetJsonFile(JsonResponses, nomSheet);
                
                    jsonFileList.AddFileToList(JsonFilePath);
                }
                else
                {
                    if (nomSheet.Contains("OXMOZ"))
                    {
                        navigatorOxmozvr.LaunchNavigatorAndGetJsonFile(url);

                        jsonFileList.AddFileToList("..\\..\\..\\..\\Json_Files\\Oxmoz.json");
                    }
                    if (nomSheet.Contains("ZEROLATENCY"))
                    {
                        navigatorZL.DownloadJsonFile(url);

                        jsonFileList.AddFileToList("..\\..\\..\\..\\Json_Files\\zerolatency.json");
                    }
                }
                
            }

            List<string> JsonFiles = jsonFileList.GetPathFile();
            int intfile = 0;

            foreach (string url in Urllist)
            {
                string nomSheet = verificationUrl.NameUrl(url);

                var sheet = excelStructure.VerificationIfFileExist(nomSheet);

                FillCellEVA fillCellEVA = new FillCellEVA(sheet);
                FillCellZLAndOx fillCellZLAndOx = new FillCellZLAndOx(sheet);


                Console.WriteLine("1");

                excelStructure.CreateStructure(path, nomSheet);

                Console.WriteLine("2");

                dataExcel.PutDataExcel(sheet);

                Console.WriteLine("3");

                Console.WriteLine("etat sauvegarde");
                excelStructure.SaveFile(path);
            }


            foreach (string url in Urllist)
            {
                string nomSheet = verificationUrl.NameUrl(url);

                var sheet = excelStructure.VerificationIfFileExist(nomSheet);

                FillCellEVA fillCellEVA = new FillCellEVA(sheet);
                FillCellZLAndOx fillCellZLAndOx = new FillCellZLAndOx(sheet);


                //Console.WriteLine("1");

                //excelStructure.CreateStructure(path, nomSheet);

                //Console.WriteLine("2");

                //dataExcel.PutDataExcel(sheet);

                //Console.WriteLine("3");

                if (nomSheet.Contains("EVA"))
                {

                    Console.WriteLine(nomSheet);

                    await takeInformationEVA.GetInformation(JsonFiles[intfile]);

                    List<string> StartedHoursEVA = takeInformationEVA.GetStartedHour();

                    List<int> MaximumPlayersEVA = takeInformationEVA.GetMaximumPlayer();

                    List<int> NumberPlayersEVA = takeInformationEVA.GetNumberPlayer();

                    List<bool> BattlepassPlayersEVA = takeInformationEVA.GetBattlePassPlayer();

                    List<bool> PeakHoursEVA = takeInformationEVA.GetPeakHour();

                    int positionX = locateCell.locateCellXPosition(date, sheet);

                    int[] positionY = locateCell.LocateCellYPosition(StartedHoursEVA, sheet, nomSheet);

                    Console.WriteLine("4");

                    fillCellEVA.FillSelectedCell(MaximumPlayersEVA, NumberPlayersEVA, sheet, positionX, positionY);

                    Console.WriteLine("5");

                    fillCellEVA.colorCell(MaximumPlayersEVA, NumberPlayersEVA, sheet, positionX, positionY, BattlepassPlayersEVA, PeakHoursEVA);

                    Console.WriteLine("6");

                    intfile++;
                }
                else
                {
                    List<string> StartedHours = new List<string>();
                    List<int> MaximumPlayers = new List<int>();
                    List<int> NumberPlayers = new List<int>();
                    List<decimal> pricesZL = new List<decimal>();

                    if (nomSheet.Contains("OXMOZ"))
                    {

                        Console.WriteLine(nomSheet);

                        JsonDocument documentOx = await takeInformationOxmozvr.GetDeserializedDocument();

                        takeInformationOxmozvr.GetInformationFromJson(documentOx);

                        StartedHours = takeInformationOxmozvr.GetStartList();

                        MaximumPlayers = takeInformationOxmozvr.GetMaximumPlayerList();

                        NumberPlayers = takeInformationOxmozvr.GetNumberPlayerList();

                        Console.WriteLine("4");
                    }
                    if (nomSheet.Contains("ZEROLATENCY"))
                    {

                        Console.WriteLine(nomSheet);

                        await takeInformationZL.ReadJson(JsonFiles[intfile]);

                        JsonDocument documentZL = takeInformationZL.DeserializeJson();

                        takeInformationZL.GetElements(documentZL);

                        StartedHours = takeInformationZL.GetHourList();

                        MaximumPlayers = takeInformationZL.GetMaximumPlayerList();

                        NumberPlayers = takeInformationZL.GetNumberPlayerList();

                        pricesZL = takeInformationZL.GetPriceList();

                        Console.WriteLine("4");
                    }

                    int positionX = locateCell.locateCellXPosition(date, sheet);

                    int[] positionY = locateCell.LocateCellYPosition(StartedHours, sheet, nomSheet);

                    fillCellZLAndOx.FillSelectedCellForZL(MaximumPlayers, NumberPlayers, sheet, positionX, positionY);

                    Console.WriteLine("5");

                    fillCellZLAndOx.ColorCell(MaximumPlayers, NumberPlayers, sheet, positionX, positionY);

                    Console.WriteLine("6");

                    intfile++;
                }
                Console.WriteLine("etat sauvegarde");
                excelStructure.SaveFile(path);
            }  
            
            
        }
    }
}
