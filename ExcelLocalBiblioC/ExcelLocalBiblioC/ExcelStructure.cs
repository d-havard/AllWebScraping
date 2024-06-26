using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronXL;
using IronXL.Formatting;
using IronXL.Styles;

namespace ExcelLocalBiblioC
{

    public class ExcelStructure
    {
        private string OriginalFilePath;
        WorkSheet workSheet;
        WorkBook workBook;

        /// <summary>
        /// The constructor for the webscraping, were you put the url of the website,
        /// the location of the original file and the location were you want to save your file 
        /// (where you have to put a @ in front of the two file location and put double \ in the two path)
        /// </summary>
        /// <param name="OriginalFilePath"></param>
        public ExcelStructure(string OriginalFilePath)
        {
            this.OriginalFilePath = OriginalFilePath;
        }

        /// <summary>
        /// Verify if the file exist in the path given. 
        /// Create an excel file with the worksheet name given if the file doesn't exist in the file,
        /// Or load the escel file if it exist.
        /// </summary>
        /// <param name="OriginalFilePath"></param>
        /// <param name="Worksheet"></param>
        public WorkSheet VerificationIfFileExist(string OriginalPath)
        {
            if (!File.Exists(OriginalFilePath))
            {
                workBook = WorkBook.Create(ExcelFileFormat.XLSX);
                workSheet = workBook.CreateWorkSheet("EVA RENNES");
            }
            else
            {
                workBook = WorkBook.LoadExcel(OriginalFilePath);
                workSheet = workBook.GetWorkSheet("EVA RENNES");
            }

            return workSheet;
        }

        /// <summary>
        /// Create the structure of the table of the Excel File.
        /// </summary>
        /// <param name="OriginalPath"></param>
        public void CreateStructure(string OriginalPath)
        {
            if (workSheet["XFD"].IsEmpty)
            {
                workSheet["XFD1"].Value = "YEP";
            }


            int a = 0;
            int minutes = 0;
            int hours = 0;
            int f = 2;
            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;
            int day = DateTime.Now.Day;
            int plusDays = 0;
            int plusMonth = 0;
            int plusYear = 0;
            int numberColumn = 0;
            ExcelFunctions excelFunctions = new ExcelFunctions(workSheet);

            string zeroBeforeTenDay = "0";
            string zeroBeforeTenMonth = "0";
            string today = $"{zeroBeforeTenDay}{day}/{zeroBeforeTenMonth}{month}/{year}";


            if (workSheet["B7"].IsEmpty)
            {
                string cell = "B7";
                workSheet["B7"].StringValue = "Heure \\ Jour";
                workSheet["B93"].StringValue = "Heure \\ Jour";
                excelFunctions.BordersThickString(cell, workSheet);
                excelFunctions.CenterTextString(cell, workSheet);


                for (int i = 0; i < 85; i++)
                {
                    minutes = a * 10;
                    a++;
                    int caseExcel = 8 + i;
                    workSheet["B" + caseExcel].FormatString = BuiltinFormats.Time3;
                    workSheet["B" + caseExcel].Value = $"{10 + hours}:{00 + minutes}";
                    workSheet["B" + caseExcel].Style.LeftBorder.Type = BorderType.Thick;
                    workSheet["B" + caseExcel].Style.RightBorder.Type = BorderType.Thick;
                    if (minutes == 50)
                    {
                        minutes++;
                        a = 0;
                        hours++;
                    }
                    if (hours == 14)
                    {
                        hours = -10;
                    }
                    workSheet["B" + caseExcel].Style.BackgroundColor = "#666666";
                    workSheet["B" + caseExcel].Style.Font.Color = "#ffffff";
                }
                workSheet["B92"].Style.BottomBorder.Type = BorderType.Thick;
                workSheet.Columns[1].Width = 9000;
            }

            if (workSheet["I4"].Style.BackgroundColor != "#f6b26b")
            {
                workSheet["I4"].Style.BackgroundColor = "#f6b26b";
                workSheet["I4"].Style.TopBorder.Type = BorderType.Thick;
                workSheet["I4"].Style.LeftBorder.Type = BorderType.Thick;
                workSheet["I4"].Style.RightBorder.Type = BorderType.Thin;
                workSheet["I4"].Style.BottomBorder.Type = BorderType.Thin;

                workSheet["I5"].Style.BackgroundColor = "#fde6cd";
                workSheet["I5"].Style.LeftBorder.Type = BorderType.Thick;
                workSheet["I5"].Style.RightBorder.Type = BorderType.Thin;
                workSheet["I5"].Style.BottomBorder.Type = BorderType.Thin;

                workSheet["I6"].Style.BackgroundColor = "#4a87e8";
                workSheet["I6"].Style.LeftBorder.Type = BorderType.Thick;
                workSheet["I6"].Style.BottomBorder.Type = BorderType.Thick;
                workSheet["I6"].Style.RightBorder.Type = BorderType.Thin;

                workSheet["J4"].StringValue = "Rempli à +75% ";
                workSheet["J4"].Style.TopBorder.Type = BorderType.Thick;
                workSheet["J4"].Style.RightBorder.Type = BorderType.Thick;
                workSheet["J4"].Style.Font.Bold = true;
                workSheet["J4"].Style.Font.Height = 14;


                workSheet["J5"].StringValue = "Rempli à +50% ";
                workSheet["J5"].Style.RightBorder.Type = BorderType.Thick;
                workSheet["J5"].Style.Font.Bold = true;
                workSheet["J5"].Style.Font.Height = 14;


                workSheet["J6"].StringValue = "Rempli en dessous de 50% ";
                workSheet["J6"].Style.RightBorder.Type = BorderType.Thick;
                workSheet["J6"].Style.BottomBorder.Type = BorderType.Thick;
                workSheet["J6"].Style.Font.Bold = true;
                workSheet["J6"].Style.Font.Height = 14;


                workSheet.Columns[9].Width = 8000;
                workSheet.Columns[8].Width = 3000;
            }
            if (workSheet["G4"].IsEmpty)
            {
                string cell = "G4";
                excelFunctions.BordersThickString(cell, workSheet);
                workSheet["G4"].Style.BottomBorder.SetColor("#00ff00");
                workSheet["G4"].Style.RightBorder.SetColor("#00ff00");
                workSheet["G4"].Value = 20;

                workSheet["H4"].Style.TopBorder.Type = BorderType.Thick;
                workSheet["H4"].Style.RightBorder.Type = BorderType.Thick;
                workSheet["H4"].Value = "HEURES CREUSES";
                workSheet["H4"].Style.Font.Bold = true;
                workSheet["H4"].Style.Font.Height = 16;

                workSheet["G5"].Style.LeftBorder.Type = BorderType.Thick;
                workSheet["G5"].Value = 23;

                workSheet["H5"].Value = "HEURES PLEINES";
                workSheet["H5"].Style.Font.Bold = true;
                workSheet["H5"].Style.Font.Height = 16;

                workSheet["G6"].Style.LeftBorder.Type = BorderType.Thick;
                workSheet["G6"].Style.BottomBorder.Type = BorderType.Thick;
                workSheet["G6"].Style.BackgroundColor = "#e06666";

                workSheet["H6"].Style.RightBorder.Type = BorderType.Thick;
                workSheet["H6"].Style.BottomBorder.Type = BorderType.Thick;
                workSheet["H6"].Value = "SESSION FULL";
                workSheet["H6"].Style.Font.Bold = true;
                workSheet["H6"].Style.Font.Height = 16;

                workSheet.Columns[7].Width = 7000;
                workSheet.Columns[6].Width = 3000;

            }


            if (workSheet["C4"].IsEmpty)
            {
                workSheet.Merge("C4:D5");
                workSheet["C4"].Value = "EVA RENNES";
                workSheet["C4"].Style.Font.Bold = true;
                workSheet["C4"].Style.Font.Height = 15;

                workSheet["C4"].Style.TopBorder.Type = BorderType.Thick;
                workSheet["D4"].Style.TopBorder.Type = BorderType.Thick;

                workSheet["C4"].Style.LeftBorder.Type = BorderType.Thick;
                workSheet["D4"].Style.RightBorder.Type = BorderType.Thick;

                workSheet["C5"].Style.LeftBorder.Type = BorderType.Thick;
                workSheet["D5"].Style.RightBorder.Type = BorderType.Thick;

                workSheet["C5"].Style.BottomBorder.Type = BorderType.Thick;
                workSheet["D5"].Style.BottomBorder.Type = BorderType.Thick;

                workSheet["C4"].Style.VerticalAlignment = VerticalAlignment.Center;
                workSheet["C4"].Style.HorizontalAlignment = HorizontalAlignment.Center;

                if (workSheet["B94"].IsEmpty)
                {
                    workSheet["B94"].Value = "Taux de remplissage journalier";
                    workSheet.Rows[93].Height = 400;
                    workSheet["B94"].Style.BackgroundColor = "#ffffff";
                    workSheet["B94"].Style.Font.Bold = true;
                    workSheet["B94"].Style.Font.Height = 10;
                }
                if (workSheet["B95"].IsEmpty)
                {
                    workSheet["B95"].Value = "Taux de remplissage hebdomadaire";
                    workSheet.Rows[94].Height = 400;
                    workSheet["B95"].Style.BackgroundColor = "#ffffff";
                    workSheet["B95"].Style.Font.Bold = true;
                    workSheet["B95"].Style.Font.Height = 10;
                }
                if (workSheet["B96"].IsEmpty)
                {
                    workSheet["B96"].Value = "Taux de remplissage mensuel";
                    workSheet.Rows[95].Height = 400;
                    workSheet["B96"].Style.BackgroundColor = "#ffffff";
                    workSheet["B96"].Style.Font.Bold = true;
                    workSheet["B96"].Style.Font.Height = 10;
                }

                if (workSheet["B99"].IsEmpty)
                {
                    workSheet["B99"].Value = "Tickets disponibles";
                    workSheet.Rows[98].Height = 400;
                    workSheet["B99"].Style.RightBorder.Type = BorderType.Thick;
                    workSheet["B99"].Style.BackgroundColor = "#ffd966";
                    workSheet["B99"].Style.VerticalAlignment = VerticalAlignment.Center;
                    workSheet["B99"].Style.HorizontalAlignment = HorizontalAlignment.Center;
                }

                if (workSheet["B100"].IsEmpty)
                {
                    workSheet["B100"].Value = "Tickets vendus";
                    workSheet.Rows[99].Height = 400;
                    workSheet["B100"].Style.RightBorder.Type = BorderType.Thick;
                    workSheet["B100"].Style.BackgroundColor = "#ff9900";
                    workSheet["B100"].Style.VerticalAlignment = VerticalAlignment.Center;
                    workSheet["B100"].Style.HorizontalAlignment = HorizontalAlignment.Center;
                }

                if (workSheet["A101"].IsEmpty)
                {
                    workSheet.Columns[0].Width = 4000;
                    workSheet["A101"].Value = "Tickets vendus / mois";
                    workSheet.Rows[100].Height = 500;
                    workSheet.Merge("A101:B101");
                    workSheet["A101"].Style.RightBorder.Type = BorderType.Thick;
                    workSheet["A101"].Style.Font.Bold = true;
                    workSheet["A101"].Style.Font.Height = 21;
                    workSheet["A101"].Style.VerticalAlignment = VerticalAlignment.Center;
                    workSheet["A101"].Style.HorizontalAlignment = HorizontalAlignment.Center;
                }

                if (workSheet["A102"].IsEmpty)
                {
                    workSheet["A102"].Value = "Taux de remplissage / mois";
                    workSheet.Rows[101].Height = 500;
                    workSheet.Merge("A102:B102");
                    workSheet["A102"].Style.Font.Bold = true;
                    workSheet["A102"].Style.Font.Height = 21;
                    workSheet["A102"].Style.VerticalAlignment = VerticalAlignment.Center;
                    workSheet["A102"].Style.HorizontalAlignment = HorizontalAlignment.Center;
                }
                if (workSheet["B103"].IsEmpty)
                {
                    workSheet["B103"].Value = "Année (en nombres)";
                    workSheet["B103"].Style.Font.Bold = false;
                    workSheet["B103"].Style.Font.Height = 11;
                    workSheet["B103"].Style.VerticalAlignment = VerticalAlignment.Center;
                    workSheet["B103"].Style.HorizontalAlignment = HorizontalAlignment.Center;
                }
                if (workSheet["A104"].IsEmpty)
                {
                    workSheet["A104"].Value = "Taux de remplissage / an";
                    workSheet.Rows[103].Height = 500;
                    workSheet.Merge("A104:B104");
                    workSheet["A104"].Style.Font.Bold = true;
                    workSheet["A104"].Style.Font.Height = 21;
                    workSheet["A104"].Style.VerticalAlignment = VerticalAlignment.Center;
                    workSheet["A104"].Style.HorizontalAlignment = HorizontalAlignment.Center;
                }
                if (workSheet["B105"].IsEmpty)
                {
                    workSheet["B105"].Value = "Année (en nombres)";
                    workSheet["B105"].Style.Font.Bold = false;
                    workSheet["B105"].Style.Font.Height = 11;
                    workSheet["B105"].Style.VerticalAlignment = VerticalAlignment.Center;
                    workSheet["B105"].Style.HorizontalAlignment = HorizontalAlignment.Center;
                }
                if (workSheet["B108"].IsEmpty)
                {
                    workSheet["B108"].Value = "Estimation CA Journalier";
                    workSheet.Rows[107].Height = 400;
                    workSheet["B108"].Style.Font.Bold = true;
                    workSheet["B108"].Style.Font.Height = 10;
                    workSheet["B108"].Style.VerticalAlignment = VerticalAlignment.Center;
                    workSheet["B108"].Style.HorizontalAlignment = HorizontalAlignment.Center;
                    workSheet.Rows[107].FormatString = BuiltinFormats.Number2;
                }
                if (workSheet["B110"].IsEmpty)
                {
                    //workSheet.Rows[109].FormatString = BuiltinFormats.Number2;
                    workSheet["B110"].StringValue = "Estimation CA Hebdomadaire";
                    workSheet.Rows[109].Height = 400;
                    workSheet["B110"].Style.Font.Bold = true;
                    workSheet["B110"].Style.Font.Height = 10;
                    workSheet["B110"].Style.VerticalAlignment = VerticalAlignment.Center;
                    workSheet["B110"].Style.HorizontalAlignment = HorizontalAlignment.Center;
                }
                if (workSheet["B112"].IsEmpty)
                {
                    //workSheet.Rows[111].FormatString = BuiltinFormats.Number2;
                    workSheet["B112"].StringValue = "Estimation CA Mensuel";
                    workSheet.Rows[111].Height = 400;
                    workSheet["B112"].Style.Font.Bold = true;
                    workSheet["B112"].Style.Font.Height = 10;
                    workSheet["B112"].Style.VerticalAlignment = VerticalAlignment.Center;
                    workSheet["B112"].Style.HorizontalAlignment = HorizontalAlignment.Center;
                }
                if (workSheet["A116"].IsEmpty)
                {
                    workSheet["A116"].Value = "CA / mois";
                    workSheet.Rows[115].Height = 500;
                    workSheet.Merge("A116:B116");
                    workSheet["A116"].Style.Font.Bold = true;
                    workSheet["A116"].Style.Font.Height = 21;
                    workSheet["A116"].Style.VerticalAlignment = VerticalAlignment.Center;
                    workSheet["A116"].Style.HorizontalAlignment = HorizontalAlignment.Center;
                }
                if (workSheet["B117"].IsEmpty)
                {
                    workSheet["B117"].Value = "Année (en nombres)";
                    workSheet["B117"].Style.Font.Bold = false;
                    workSheet["B117"].Style.Font.Height = 11;
                    workSheet["B117"].Style.VerticalAlignment = VerticalAlignment.Center;
                    workSheet["B117"].Style.HorizontalAlignment = HorizontalAlignment.Center;
                }
                if (workSheet["A118"].IsEmpty)
                {
                    workSheet["A118"].Value = "CA / an";
                    workSheet.Rows[115].Height = 500;
                    workSheet.Merge("A118:B118");
                    workSheet["A118"].Style.Font.Bold = true;   
                    workSheet["A118"].Style.Font.Height = 21;
                    workSheet["A118"].Style.VerticalAlignment = VerticalAlignment.Center;
                    workSheet["A118"].Style.HorizontalAlignment = HorizontalAlignment.Center;
                }
                if (workSheet["B119"].IsEmpty)
                {
                    workSheet["B119"].Value = "Année (en nombres)";
                    workSheet["B119"].Style.Font.Bold = false;
                    workSheet["B119"].Style.Font.Height = 11;
                    workSheet["B119"].Style.VerticalAlignment = VerticalAlignment.Center;
                    workSheet["B119"].Style.HorizontalAlignment = HorizontalAlignment.Center;
                }
                if (workSheet["B120"].IsEmpty)
                {
                    string cell = "B120";
                    workSheet["B120"].StringValue = "8/8";
                    workSheet["B120"].Style.BackgroundColor = "#e06666";
                    excelFunctions.BordersThinString(cell, workSheet);
                    excelFunctions.CenterTextString(cell, workSheet);

                    cell = "B121";
                    workSheet["B121"].StringValue = "7/8";
                    workSheet["B121"].Style.BackgroundColor = "#f6b26b";
                    excelFunctions.BordersThinString(cell, workSheet);
                    excelFunctions.CenterTextString(cell, workSheet);

                    cell = "B122";
                    workSheet["B122"].StringValue = "6/8";
                    workSheet["B122"].Style.BackgroundColor = "#f6b26b";
                    excelFunctions.BordersThinString(cell, workSheet);
                    excelFunctions.CenterTextString(cell, workSheet);

                    cell = "B123";
                    workSheet["B123"].StringValue = "5/8";
                    workSheet["B123"].Style.BackgroundColor = "#fde6cd";
                    excelFunctions.BordersThinString(cell, workSheet);
                    excelFunctions.CenterTextString(cell, workSheet);

                    cell = "B124";
                    workSheet["B124"].StringValue = "4/8";
                    workSheet["B124"].Style.BackgroundColor = "#4a87e8";
                    excelFunctions.BordersThinString(cell, workSheet);
                    excelFunctions.CenterTextString(cell, workSheet);

                    cell = "B125";
                    workSheet["B125"].StringValue = "3/8";
                    workSheet["B125"].Style.BackgroundColor = "#4a87e8";
                    excelFunctions.BordersThinString(cell, workSheet);
                    excelFunctions.CenterTextString(cell, workSheet);

                    cell = "B126";
                    workSheet["B126"].StringValue = "2/8";
                    workSheet["B126"].Style.BackgroundColor = "#4a87e8";
                    excelFunctions.BordersThinString(cell, workSheet);
                    excelFunctions.CenterTextString(cell, workSheet);

                    cell = "B127";
                    workSheet["B127"].StringValue = "1/8";
                    workSheet["B127"].Style.BackgroundColor = "#4a87e8";
                    excelFunctions.BordersThinString(cell, workSheet);
                    excelFunctions.CenterTextString(cell, workSheet);
                }
            }

            ////workSheet.AutoSizeColumn(1, false);
            IronXL.Range range2 = workSheet["C7:ZZ7"];
            // Set the data format to 1/1/2020 12:12:12
            range2.FormatString = BuiltinFormats.LongDate1;
        }

        public void SaveFile()
        {
            workBook.SaveAs("C:\\Stage\\AllWebScraping_VirtualGame\\dataFormatts.xlsx");
        }
    }
}