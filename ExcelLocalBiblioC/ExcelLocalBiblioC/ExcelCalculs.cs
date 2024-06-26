using IronXL;
using IronXL.Formatting;
using IronXL.Styles;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLocalBiblioC
{
    public class ExcelCalculs
    {
        WorkSheet workSheet;
        ExcelFunctions excelFunctions;
        PutTimeCells timeFunctions;
       
        public ExcelCalculs(WorkSheet workSheet) 
        {
            
            this.workSheet = workSheet;
        }

        public void CalculateTickets(int columnWhile)
        {
            ExcelFunctions excelFunctions = new ExcelFunctions(workSheet);
            PutTimeCells timeFunctions = new PutTimeCells(workSheet);
            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;
            int day = DateTime.Now.Day;
            string zeroBeforeTenDay = "0";
            string zeroBeforeTenMonth = "0";
            string todayForTickets = $"{zeroBeforeTenDay}{day}/{zeroBeforeTenMonth}{month}/{year}";

            if (day < 10 && month < 10)
            {
                todayForTickets = $"{zeroBeforeTenMonth}{day}/{zeroBeforeTenMonth}{month}/{year}";
            }
            if (day < 10 && month >= 10)
            {
                todayForTickets = $"{zeroBeforeTenDay}{day}/{month}/{year}";
            }
            if (day >= 10 && month < 10)
            {
                todayForTickets = $"{day}/{zeroBeforeTenMonth}{month}/{year}";
            }
            if (day >= 10 && month >= 10)
            {
                todayForTickets = $"{day}/{month}/{year}";
            }
            decimal ticketsVendus = 0;
            decimal ticketsDispos = 0;
            int columnWhileTickets = 2;
            bool calculatedTickets = false;
            DateTime currentDay = DateTime.Today;
            bool monthTRPut = false;
            // ! Attention ! Les valeurs des colonnes et des lignes commencent à 0 quand ce sont des nombres !
            //Boucle pour calculer les tickets vendus
            while (!calculatedTickets)
            {
                string[] cellDay = workSheet.Rows[6].Columns[columnWhileTickets].Value.ToString().Split(' ');
                //On split les espaces dans la date puis on regarde si le jour sélectionner est égal à aujourd'hui puis on regarde toutes
                //les cellules du jour pour calculer les tickets vendus.
                if (cellDay[0] == todayForTickets)
                {
                    for (int i = 8; i < 92; i++)
                    {
                        if (workSheet.Rows[i].Columns[columnWhileTickets].IsEmpty)
                        {

                        }
                        else
                        {
                            if (workSheet.Rows[i].Columns[columnWhileTickets].ToString().Contains('/'))
                            {
                                string[] tickets = workSheet.Rows[i].Columns[columnWhileTickets].ToString().Split('/');
                                ticketsVendus += Convert.ToDecimal(tickets[0]);
                                ticketsDispos += Convert.ToDecimal(tickets[1]);
                            }

                        }
                    }
                    //On regarde si les tickets disponible sont supérieur à 0. Car en cas de jour fermé
                    //ou alors de problème la calcul ne se fera pas.
                    //Puis on calcule le taux de remplissage de la journée.
                    if (ticketsDispos > 0)
                    {
                        workSheet.Rows[93].Columns[columnWhileTickets].FormatString = BuiltinFormats.Percent2;
                        decimal tauxRemplissageJourna = ticketsVendus / ticketsDispos;
                        workSheet.Rows[93].Columns[columnWhileTickets].Value = tauxRemplissageJourna;
                        workSheet.Rows[93].Columns[columnWhileTickets].Style.Font.Height = 15;
                        workSheet.Rows[93].Columns[columnWhileTickets].Style.Font.Bold = true;
                        excelFunctions.BodersThinInt(93, columnWhileTickets, workSheet);
                        excelFunctions.PutBordersBlackInt(93, columnWhile, workSheet);
                        excelFunctions.CenterTextInt(93, columnWhileTickets, workSheet);
                    }

                    //On met les valeurs obtenue dans les bonnes cases
                    workSheet.Rows[98].Columns[columnWhileTickets].Value = ticketsDispos;
                    workSheet.Rows[98].Columns[columnWhileTickets].Style.Font.Height = 20;
                    workSheet.Rows[98].Columns[columnWhileTickets].Style.Font.Bold = true;
                    excelFunctions.BodersThinInt(98, columnWhileTickets, workSheet);
                    excelFunctions.CenterTextInt(98, columnWhileTickets, workSheet);
                    excelFunctions.PutBordersBlackInt(98, columnWhile, workSheet);


                    workSheet.Rows[99].Columns[columnWhileTickets].Value = ticketsVendus;
                    workSheet.Rows[99].Columns[columnWhileTickets].Style.Font.Height = 20;
                    workSheet.Rows[99].Columns[columnWhileTickets].Style.Font.Bold = true;
                    excelFunctions.BodersThinInt(99, columnWhileTickets, workSheet);
                    excelFunctions.CenterTextInt(99, columnWhileTickets, workSheet);
                    excelFunctions.PutBordersBlackInt(99, columnWhile, workSheet);


                    var cellValue = workSheet.Rows[99].Columns[columnWhileTickets].Value;
                    decimal numberSelledTickets = Convert.ToDecimal(cellValue);
                    decimal CAperDay = 0;
                    decimal ticketsInCell = 0;

                    //On regarde si la bordure est verte pour savoir si c'est en heure creuse ou pleine
                    //Puis on multiplie par la bonne valeur, qui est le prix par ticket
                    for (int i = 8; i < 92; i++)
                    {
                        if (workSheet.Rows[i].Columns[columnWhileTickets].IsEmpty)
                        {

                        }
                        else 
                        {
                            if (workSheet.Rows[i].Columns[columnWhileTickets].Style.RightBorder.Color == "#00ff02")
                            {
                                if (workSheet.Rows[i].Columns[columnWhileTickets].ToString().Contains('/'))
                                {
                                    string[] ticketsInCellSplit = workSheet.Rows[i].Columns[columnWhileTickets].ToString().Split('/');
                                    ticketsInCell = Convert.ToDecimal(ticketsInCellSplit[0]);
                                    CAperDay += (ticketsInCell * 20);
                                }
                                
                            }
                            else
                            {
                                if (workSheet.Rows[i].Columns[columnWhileTickets].ToString().Contains('/'))
                                {
                                    string[] ticketsInCellSplit = workSheet.Rows[i].Columns[columnWhileTickets].ToString().Split('/');
                                    ticketsInCell = Convert.ToDecimal(ticketsInCellSplit[0]);
                                    CAperDay += (ticketsInCell * 23);
                                }
                            }
                            
                        }
                    }

                    //Puis on met les valeurs dans les cases
                    workSheet.Rows[107].Columns[columnWhileTickets].Value = CAperDay;
                    workSheet.Rows[107].Columns[columnWhileTickets].Style.Font.Height = 14;
                    workSheet.Rows[107].Columns[columnWhileTickets].Style.Font.Bold = true;
                    excelFunctions.BodersThinInt(107, columnWhileTickets, workSheet);
                    excelFunctions.CenterTextInt(107, columnWhileTickets, workSheet);
                    excelFunctions.PutBordersBlackInt(107, columnWhile, workSheet);

                    bool calculatedTicketsWeek = false;
                    while (!calculatedTicketsWeek)
                    {
                        string dayForCalculateCA = currentDay.DayOfWeek.ToString();
                        if (cellDay[0] == todayForTickets)
                        {
                            switch (dayForCalculateCA)
                            {
                                case "Monday":

                                    workSheet.Rows[94].Columns[columnWhileTickets].FormatString = BuiltinFormats.Percent2;
                                    if (ticketsDispos > 0)
                                    {
                                        workSheet.Rows[94].Columns[columnWhileTickets].Value = ticketsVendus / ticketsDispos;
                                    }
                                    workSheet.Rows[94].Columns[columnWhileTickets].Style.Font.Height = 15;
                                    workSheet.Rows[94].Columns[columnWhileTickets].Style.Font.Bold = true;
                                    excelFunctions.BodersThinInt(94, columnWhileTickets, workSheet);
                                    excelFunctions.PutBordersBlackInt(94, columnWhile, workSheet);
                                    excelFunctions.CenterTextInt(94, columnWhileTickets, workSheet);

                                    workSheet.Rows[109].Columns[columnWhileTickets].Value = workSheet.Rows[107].Columns[columnWhileTickets];
                                    workSheet.Rows[109].Columns[columnWhileTickets].Style.Font.Height = 14;
                                    workSheet.Rows[109].Columns[columnWhileTickets].Style.Font.Bold = true;
                                    workSheet.Rows[109].Columns[columnWhileTickets].Style.VerticalAlignment = VerticalAlignment.Center;
                                    excelFunctions.BodersThinInt(109, columnWhileTickets, workSheet);
                                    excelFunctions.CenterTextInt(109, columnWhileTickets, workSheet);
                                    excelFunctions.PutBordersBlackInt(109, columnWhile, workSheet);

                                    break;

                                case "Tuesday":
                                    int columnMergedTuesday = columnWhileTickets - 1;
                                    var cellAdress1TuesdaySum = workSheet.Rows[107].Columns[columnWhileTickets - 1].RangeAddressAsString;
                                    var cellAdress2TuesdaySum = workSheet.Rows[107].Columns[columnWhileTickets].RangeAddressAsString;
                                    var rangeTuesdayCA = workSheet[$"{cellAdress1TuesdaySum}:{cellAdress2TuesdaySum}"];

                                    var cellAdress1TuesdayAvg = workSheet.Rows[93].Columns[columnWhileTickets - 1].RangeAddressAsString;
                                    var cellAdress2TuesdayAvg = workSheet.Rows[93].Columns[columnWhileTickets].RangeAddressAsString;
                                    var rangeTuesdayAvg = workSheet[$"{cellAdress1TuesdayAvg}:{cellAdress2TuesdayAvg}"];

                                    workSheet.Rows[109].Columns[columnMergedTuesday].Value = rangeTuesdayCA.Sum();
                                    workSheet.Rows[94].Columns[columnMergedTuesday].Value = rangeTuesdayAvg.Avg();
                                    break;

                                case "Wednesday":
                                    int columnMergedWednesday = columnWhileTickets - 2;
                                    var cellAdress1WednesdaySum = workSheet.Rows[107].Columns[columnWhileTickets - 2].RangeAddressAsString;
                                    var cellAdress2WednesdaySum = workSheet.Rows[107].Columns[columnWhileTickets].RangeAddressAsString;
                                    var rangeWednesdayCA = workSheet[$"{cellAdress1WednesdaySum}:{cellAdress2WednesdaySum}"];

                                    var cellAdress1WednesdayAvg = workSheet.Rows[93].Columns[columnWhileTickets - 2].RangeAddressAsString;
                                    var cellAdress2WednesdayAvg = workSheet.Rows[93].Columns[columnWhileTickets].RangeAddressAsString;
                                    var rangeWednesdayAvg = workSheet[$"{cellAdress1WednesdayAvg}:{cellAdress2WednesdayAvg}"];

                                    workSheet.Rows[109].Columns[columnMergedWednesday].Value = rangeWednesdayCA.Sum();
                                    workSheet.Rows[94].Columns[columnMergedWednesday].Value = rangeWednesdayAvg.Avg();
                                    break;

                                case "Thursday":
                                    int columnMergedThursday = columnWhileTickets - 3;
                                    var cellAdress1ThursdaySum = workSheet.Rows[107].Columns[columnWhileTickets - 3].RangeAddressAsString;
                                    var cellAdress2ThursdaySum = workSheet.Rows[107].Columns[columnWhileTickets].RangeAddressAsString;
                                    var rangeThursdayCA = workSheet[$"{cellAdress1ThursdaySum}:{cellAdress2ThursdaySum}"];

                                    var cellAdress1ThursdayAvg = workSheet.Rows[93].Columns[columnWhileTickets - 3].RangeAddressAsString;
                                    var cellAdress2ThursdayAvg = workSheet.Rows[93].Columns[columnWhileTickets].RangeAddressAsString;
                                    var rangeThursdayAvg = workSheet[$"{cellAdress1ThursdayAvg}:{cellAdress2ThursdayAvg}"];

                                    workSheet.Rows[109].Columns[columnMergedThursday].Value = rangeThursdayCA.Sum();
                                    workSheet.Rows[94].Columns[columnMergedThursday].Value = rangeThursdayAvg.Avg();
                                    break;

                                case "Friday":
                                    int columnMergedFriday = columnWhileTickets - 4;
                                    var cellAdress1FridaySum = workSheet.Rows[107].Columns[columnWhileTickets - 4].RangeAddressAsString;
                                    var cellAdress2FridaySum = workSheet.Rows[107].Columns[columnWhileTickets].RangeAddressAsString;
                                    var rangeFridayCA = workSheet[$"{cellAdress1FridaySum}:{cellAdress2FridaySum}"];

                                    var cellAdress1FridayAvg = workSheet.Rows[93].Columns[columnWhile - 4].RangeAddressAsString;
                                    var cellAdress2FridayAvg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                    var rangeFridayAvg = workSheet[$"{cellAdress1FridayAvg}:{cellAdress2FridayAvg}"];

                                    workSheet.Rows[109].Columns[columnMergedFriday].Value = rangeFridayCA.Sum();
                                    workSheet.Rows[94].Columns[columnMergedFriday].Value = rangeFridayAvg.Avg();
                                    break;

                                case "Saturday":
                                    int columnMergedSaturday = columnWhileTickets - 5;
                                    var cellAdress1SaturdaySum = workSheet.Rows[107].Columns[columnWhileTickets - 5].RangeAddressAsString;
                                    var cellAdress2SaturdaySum = workSheet.Rows[107].Columns[columnWhileTickets].RangeAddressAsString;
                                    var rangeSaturdayCA = workSheet[$"{cellAdress1SaturdaySum}:{cellAdress2SaturdaySum}"];

                                    var cellAdress1SaturdayAvg = workSheet.Rows[93].Columns[columnWhile - 5].RangeAddressAsString;
                                    var cellAdress2SaturdayAvg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                    var rangeSaturdayAvg = workSheet[$"{cellAdress1SaturdayAvg}:{cellAdress2SaturdayAvg}"];

                                    workSheet.Rows[109].Columns[columnMergedSaturday].Value = rangeSaturdayCA.Sum();
                                    workSheet.Rows[94].Columns[columnMergedSaturday].Value = rangeSaturdayAvg.Avg();
                                    break;

                                case "Sunday":
                                    int columnMergedSunday = columnWhileTickets - 6;
                                    var cellAdress1SundaySum = workSheet.Rows[107].Columns[columnWhileTickets - 6].RangeAddressAsString;
                                    var cellAdress2SundaySum = workSheet.Rows[107].Columns[columnWhileTickets].RangeAddressAsString;
                                    var rangeSundayCA = workSheet[$"{cellAdress1SundaySum}:{cellAdress2SundaySum}"];

                                    var cellAdress1SundayAvg = workSheet.Rows[93].Columns[columnWhile - 6].RangeAddressAsString;
                                    var cellAdress2SundayAvg = workSheet.Rows[93].Columns[columnWhile - 6].RangeAddressAsString;
                                    var rangeSundayAvg = workSheet[$"{cellAdress1SundayAvg}:{cellAdress2SundayAvg}"];

                                    workSheet.Rows[109].Columns[columnMergedSunday].Value = rangeSundayCA.Sum();
                                    workSheet.Rows[94].Columns[columnMergedSunday].Value = rangeSundayAvg.Avg();
                                    break;
                            }

                            //On initialise les variables et les "range" utilisé plus tard
                            DateTime date = DateTime.Now;
                            string[] actualDayMonth = date.ToString().Split('/');
                            var firstDayOfMonth = new DateTime(date.Year, date.Month, 1);
                            var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
                            string[] firstDayMonth = firstDayOfMonth.ToString().Split('/');
                            string[] lastDayMonth = lastDayOfMonth.ToString().Split('/');
                            int dayMonthInt = Int32.Parse(firstDayMonth[0]);
                            int lastDayMonthInt = Int32.Parse(lastDayMonth[0]);
                            int actualDayMonthInt = Int32.Parse(actualDayMonth[0]);
                            Console.WriteLine(dayMonthInt);
                            Console.WriteLine(lastDayMonthInt);
                            Console.WriteLine(firstDayOfMonth.ToString() + " / " + lastDayOfMonth.ToString());
                            var cellForRangeTR1 = workSheet.Rows[94].Columns[columnWhileTickets - actualDayMonthInt + 1].RangeAddressAsString;
                            var cellForRangeTR2 = workSheet.Rows[94].Columns[columnWhileTickets - actualDayMonthInt + lastDayMonthInt].RangeAddressAsString;
                            var rangeMonthlyTR = workSheet[$"{cellForRangeTR1}:{cellForRangeTR2}"];
                            var cellForRangeCA1 = workSheet.Rows[109].Columns[columnWhileTickets - actualDayMonthInt + 1].RangeAddressAsString;
                            var cellForRangecA2 = workSheet.Rows[109].Columns[columnWhileTickets - actualDayMonthInt + lastDayMonthInt].RangeAddressAsString;
                            var rangeMonthlyCA = workSheet[$"{cellForRangeCA1}:{cellForRangecA2}"];
                            var cellForRangeSelledTickets1 = workSheet.Rows[99].Columns[columnWhileTickets - actualDayMonthInt + 1].RangeAddressAsString;
                            var cellForRangeSelledTickets2 = workSheet.Rows[99].Columns[columnWhileTickets - actualDayMonthInt + lastDayMonthInt].RangeAddressAsString;
                            var rangeMonthlySelledTickets = workSheet[$"{cellForRangeSelledTickets1}:{cellForRangeSelledTickets2}"];
                            workSheet.Rows[95].Columns[columnWhileTickets - actualDayMonthInt + 1].Value = rangeMonthlyTR.Avg();
                            workSheet.Rows[111].Columns[columnWhileTickets - actualDayMonthInt + 1].Value = rangeMonthlyCA.Sum();
                            workSheet.Rows[95].Columns[columnWhileTickets - actualDayMonthInt + 1].FormatString = BuiltinFormats.Percent2;
                            var SelledTicketsPerMonth = rangeMonthlySelledTickets.Sum();

                            
                            int columnWhilePutTR = 2;
                            bool putTRMonthly = false;
                            //On calcule le Taux de remplissage /mois et /an
                            while (!putTRMonthly)
                            {
                                workSheet.Rows[102].Columns[columnWhilePutTR].FormatString = BuiltinFormats.Accounting0;
                                var actualMonth = date.Month;
                                var actualYear = date.Year;
                                if (workSheet.Rows[101].Columns[columnWhilePutTR].IsEmpty && workSheet.Rows[102].Columns[columnWhilePutTR].IsEmpty && workSheet.Rows[104].Columns[columnWhilePutTR].IsEmpty)
                                {

                                    workSheet.Rows[102].Columns[columnWhilePutTR].Value = date.Month;
                                    workSheet.Rows[104].Columns[columnWhilePutTR].Value = date.Year;
                                    workSheet.Rows[101].Columns[columnWhilePutTR].FormatString = BuiltinFormats.Percent2;
                                    workSheet.Rows[101].Columns[columnWhilePutTR].Value = rangeMonthlyTR.Avg();
                                    excelFunctions.CenterTextInt(100, columnWhilePutTR, workSheet);
                                    excelFunctions.CenterTextInt(101, columnWhilePutTR, workSheet);

                                    excelFunctions.CenterTextInt(104, columnWhilePutTR, workSheet);
                                    workSheet.Rows[100].Columns[columnWhilePutTR].Value = SelledTicketsPerMonth;
                                    workSheet.Rows[101].Columns[columnWhilePutTR].Style.Font.Height = 18;
                                    workSheet.Rows[101].Columns[columnWhilePutTR].Style.Font.Bold = true;
                                    workSheet.Rows[100].Columns[columnWhilePutTR].Style.Font.Height = 18;
                                    workSheet.Rows[100].Columns[columnWhilePutTR].Style.Font.Bold = true;

                                    putTRMonthly = true;
                                    
                                }
                                else
                                {
                                    var cellMonthObject = workSheet.Rows[102].Columns[columnWhilePutTR].Value;
                                    int cellMonthInt = Int32.Parse(cellMonthObject.ToString());
                                    var cellYearObject = workSheet.Rows[104].Columns[columnWhilePutTR].Value;
                                    int cellYearInt = Int32.Parse(cellYearObject.ToString());
                                    var cellYearAvg1 = workSheet.Rows[101].Columns[columnWhilePutTR - cellMonthInt].RangeAddressAsString;
                                    var cellYearAvg2 = workSheet.Rows[101].Columns[columnWhilePutTR].RangeAddressAsString;
                                    var rangeYearAvg = workSheet[$"{cellYearAvg1}:{cellYearAvg2}"];
                                    if (cellMonthInt == actualMonth && cellYearInt == actualYear)
                                    {
                                        workSheet.Rows[101].Columns[columnWhilePutTR].FormatString = BuiltinFormats.Percent2;
                                        workSheet.Rows[101].Columns[columnWhilePutTR].Value = rangeMonthlyTR.Avg();
                                        workSheet.Rows[103].Columns[columnWhilePutTR - actualMonth + 1].FormatString = BuiltinFormats.Percent2;
                                        workSheet.Rows[103].Columns[columnWhilePutTR - actualMonth + 1].Value = rangeYearAvg.Avg();
                                        workSheet.Rows[103].Columns[columnWhilePutTR - actualMonth + 1].Style.Font.Height = 18;
                                        workSheet.Rows[103].Columns[columnWhilePutTR - actualMonth + 1].Style.Font.Bold = true;
                                        excelFunctions.CenterTextInt(103, columnWhilePutTR - actualMonth + 1, workSheet);
                                        putTRMonthly = true;
                                    }
                                }
                                columnWhilePutTR++;
                            }


                            int columnWhilePutCA = 2;
                            bool putCAMonthly = false;

                            //On calcule le CA /mois et /an
                            while (!putCAMonthly)
                            {
                                var actualMonth = date.Month;
                                var actualYear = date.Year;
                                if (workSheet.Rows[115].Columns[columnWhilePutCA].IsEmpty && workSheet.Rows[116].Columns[columnWhilePutCA].IsEmpty && workSheet.Rows[118].Columns[columnWhilePutCA].IsEmpty)
                                {

                                    workSheet.Rows[116].Columns[columnWhilePutCA].Value = date.Month;
                                    workSheet.Rows[118].Columns[columnWhilePutCA].Value = date.Year;
                                    workSheet.Rows[115].Columns[columnWhilePutCA].FormatString = BuiltinFormats.Accounting0;
                                    workSheet.Rows[115].Columns[columnWhilePutCA].Value = rangeMonthlyCA.Sum();
                                    excelFunctions.CenterTextInt(115, columnWhilePutCA, workSheet);
                                    excelFunctions.CenterTextInt(117, columnWhilePutCA, workSheet);

                                    excelFunctions.CenterTextInt(115, columnWhilePutCA, workSheet);
                                    workSheet.Rows[115].Columns[columnWhilePutCA].Style.Font.Height = 18;
                                    workSheet.Rows[115].Columns[columnWhilePutCA].Style.Font.Bold = true;
                                    workSheet.Rows[117].Columns[columnWhilePutCA].Style.Font.Height = 18;
                                    workSheet.Rows[117].Columns[columnWhilePutCA].Style.Font.Bold = true;

                                    putCAMonthly = true;

                                }
                                else
                                {
                                    var cellMonthObject = workSheet.Rows[102].Columns[columnWhilePutCA].Value;
                                    int cellMonthInt = Int32.Parse(cellMonthObject.ToString());
                                    var cellYearObject = workSheet.Rows[104].Columns[columnWhilePutCA].Value;
                                    int cellYearInt = Int32.Parse(cellYearObject.ToString());
                                    var cellYearSum1 = workSheet.Rows[115].Columns[columnWhilePutCA - cellMonthInt].RangeAddressAsString;
                                    var cellYearSum2 = workSheet.Rows[115].Columns[columnWhilePutCA].RangeAddressAsString;
                                    var rangeYearSum = workSheet[$"{cellYearSum1}:{cellYearSum2}"];
                                    if (cellMonthInt == actualMonth && cellYearInt == actualYear)
                                    {
                                        workSheet.Rows[115].Columns[columnWhilePutCA].FormatString = BuiltinFormats.Accounting0;
                                        workSheet.Rows[115].Columns[columnWhilePutCA].Value = rangeMonthlyCA.Sum();
                                        workSheet.Rows[117].Columns[columnWhilePutCA - actualMonth + 1].FormatString = BuiltinFormats.Accounting0;
                                        workSheet.Rows[117].Columns[columnWhilePutCA - actualMonth + 1].Value = rangeYearSum.Sum();
                                        workSheet.Rows[117].Columns[columnWhilePutCA - actualMonth + 1].Style.Font.Height = 18;
                                        workSheet.Rows[117].Columns[columnWhilePutCA - actualMonth + 1].Style.Font.Bold = true;
                                        excelFunctions.CenterTextInt(117, columnWhilePutCA - actualMonth + 1, workSheet);
                                        putCAMonthly = true;
                                    }
                                }
                                columnWhilePutCA++;
                            }
                            calculatedTicketsWeek = true;
                        }
                    }
                        

                    calculatedTickets = true;
                }
                columnWhileTickets++;
            }
        }
    }
}
