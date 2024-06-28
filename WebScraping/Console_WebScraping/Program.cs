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

using IronXL.Formatting;
using IronXL.Styles;
using NUnit.Framework.Internal.Execution;
using Microsoft.VisualBasic;
using NUnit.Framework.Constraints;

namespace Console_WebScraping
{
    internal class Program
    {
        static async Task Main(string[] args)
        {

            using PeriodicTimer timer = new(TimeSpan.FromMilliseconds(60000));

            UrlList urlList = new UrlList();

            urlList.MakeUrlList();

            List<string> Urllist = urlList.GetUrlList();
            string path = "E:\\Stage\\Virtual_game\\WebScraping\\dataFormatts.xlsx";
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
            string date = DateTime.Now.ToString("dd/MM/yyyy");

            string JsonFilePath = "";
            while (true)
            {
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
                        if (nomSheet.Contains("ZEROLATENCY"))
                        {
                            navigatorZL.DownloadJsonFile(url, nomSheet);




                            jsonFileList.AddFileToList($"{nomSheet}.json");
                        }
                    }

                }

                List<string> JsonFiles = jsonFileList.GetPathFile();
                int intfile = 0;
                int dataExcelReturn = 0;


                foreach (string url in Urllist)
                {
                    string nomSheet = verificationUrl.NameUrl(url);

                    var sheet = excelStructure.VerificationIfFileExist(nomSheet);



                    



                    FillCellEVA fillCellEVA = new FillCellEVA(sheet);
                    FillCellZLAndOx fillCellZLAndOx = new FillCellZLAndOx(sheet);


                    Console.WriteLine($"Create structure {nomSheet}");

                    excelStructure.CreateStructure(path, nomSheet);

                    Console.WriteLine($"Put Data in the Excel in the sheet {nomSheet}");

                    dataExcelReturn = dataExcel.PutDataExcel(sheet);

                    

                    Console.WriteLine("etat sauvegarde");
                    excelStructure.SaveFile();
                }

                int priceZLplace = 0;
                foreach (string url in Urllist)
                {
                    string nomSheet = verificationUrl.NameUrl(url);

                    var sheet = excelStructure.VerificationIfFileExist(nomSheet);

                    string today = DateTime.Now.ToString("dddd");

                    FillCellEVA fillCellEVA = new FillCellEVA(sheet);
                    FillCellZLAndOx fillCellZLAndOx = new FillCellZLAndOx(sheet);
                    PutTimeCells putTimeCells = new PutTimeCells(sheet);
                    ExcelCalculs excelCalculs = new ExcelCalculs(sheet);


                    //Console.WriteLine("1");

                    //excelStructure.CreateStructure(path, nomSheet);

                    //Console.WriteLine("2");

                    //dataExcel.PutDataExcel(sheet);

                    //Console.WriteLine("3");

                    if (nomSheet.Contains("EVA"))
                    {

                        Console.WriteLine(nomSheet);

                        Console.WriteLine($"Get Information {nomSheet}");

                        await takeInformationEVA.GetInformation(JsonFiles[intfile]);

                        


                        List<string> StartedHoursEVA = takeInformationEVA.GetStartedHour();

                        List<int> MaximumPlayersEVA = takeInformationEVA.GetMaximumPlayer();

                        List<int> NumberPlayersEVA = takeInformationEVA.GetNumberPlayer();

                        List<bool> BattlepassPlayersEVA = takeInformationEVA.GetBattlePassPlayer();

                        List<bool> PeakHoursEVA = takeInformationEVA.GetPeakHour();

                        Console.WriteLine($"Locate the X position in the sheet {nomSheet}");

                        int positionX = locateCell.locateCellXPosition(date, sheet);

                        Console.WriteLine($"Locate the Y position in the sheet {nomSheet}");

                        int[] positionY = locateCell.LocateCellYPosition(StartedHoursEVA, sheet, nomSheet);

                        Console.WriteLine($"Fill all the selected cells in the sheet {nomSheet}");

                        fillCellEVA.FillSelectedCell(MaximumPlayersEVA, NumberPlayersEVA, sheet, positionX, positionY);


                        Console.WriteLine($"Color the selected cells in the sheet {nomSheet}");


                        fillCellEVA.colorCell(MaximumPlayersEVA, NumberPlayersEVA, sheet, positionX, positionY, BattlepassPlayersEVA, PeakHoursEVA);


                        Console.WriteLine($"Calculate the tickets in the sheet {nomSheet}");

                        excelCalculs.CalculateTickets(dataExcelReturn, nomSheet);


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

                            Console.WriteLine($"Get Information {nomSheet}");

                            JsonDocument documentOx = await takeInformationOxmozvr.GetDeserializedDocument();

                            takeInformationOxmozvr.GetInformationFromJson(documentOx);

                            StartedHours = takeInformationOxmozvr.GetStartList();

                            MaximumPlayers = takeInformationOxmozvr.GetMaximumPlayerList();

                            NumberPlayers = takeInformationOxmozvr.GetNumberPlayerList();

                            
                        }
                        if (nomSheet.Contains("ZEROLATENCY"))
                        {

                            Console.WriteLine(nomSheet);

                            Console.WriteLine($"Get Information {nomSheet}");

                            await takeInformationZL.ReadJson(JsonFiles[intfile]);

                            JsonDocument documentZL = takeInformationZL.DeserializeJson();

                            takeInformationZL.GetElements(documentZL);

                            StartedHours = takeInformationZL.GetHourList();

                            MaximumPlayers = takeInformationZL.GetMaximumPlayerList();

                            NumberPlayers = takeInformationZL.GetNumberPlayerList();

                            pricesZL = takeInformationZL.GetPriceList();
                        }

                        Console.WriteLine($"Locate the X position in the sheet {nomSheet}");

                        int positionX = locateCell.locateCellXPosition(date, sheet);

                        Console.WriteLine($"Locate the Y position in the sheet {nomSheet}");

                        int[] positionY = locateCell.LocateCellYPosition(StartedHours, sheet, nomSheet);

                        Console.WriteLine($"Fill all the selected cells in the sheet {nomSheet}");

                        fillCellZLAndOx.FillSelectedCellForZL(MaximumPlayers, NumberPlayers, sheet, positionX, positionY);

                        Console.WriteLine($"Color the selected cells in the sheet {nomSheet}");

                        fillCellZLAndOx.ColorCell(MaximumPlayers, NumberPlayers, sheet, positionX, positionY);

                        Console.WriteLine($"Calculate the tickets in the sheet {nomSheet}");

                        excelCalculs.CalculateTickets(dataExcelReturn, nomSheet, pricesZL[priceZLplace]);

                        intfile++;

                        priceZLplace++;
                    }
                    Console.WriteLine("etat sauvegarde");
                    excelStructure.SaveFile();
                }

                Console.WriteLine("done");
            }
            
            


           
        }
    }
}
