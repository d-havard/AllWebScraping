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
        DateTime currentDay;
        public ExcelCalculs(WorkSheet workSheet) 
        {
            this.workSheet = workSheet;
            this.excelFunctions = new ExcelFunctions(workSheet);
            DateTime currentDay = DateTime.Today;
        }

        public void CalculateTickets(string today, int columnWhile)
        {

            decimal ticketsVendus = 0;
            decimal ticketsDispos = 0;
            int columnWhileTickets = 2;
            bool calculatedTickets = false;

            while (!calculatedTickets)
            {
                string[] cellDay = workSheet.Rows[6].Columns[columnWhileTickets].Value.ToString().Split(' ');
                if (cellDay[0] == today)
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
                    if (ticketsDispos > 0)
                    {
                        workSheet.Rows[93].Columns[columnWhileTickets].FormatString = BuiltinFormats.Percent2;
                        workSheet.Rows[93].Columns[columnWhileTickets].Value = ticketsVendus / ticketsDispos;
                        workSheet.Rows[93].Columns[columnWhileTickets].Style.Font.Height = 15;
                        workSheet.Rows[93].Columns[columnWhileTickets].Style.Font.Bold = true;
                        excelFunctions.BodersThinInt(93, columnWhileTickets, workSheet);
                        excelFunctions.PutBordersBlackInt(93, columnWhile, workSheet);
                        excelFunctions.CenterTextInt(93, columnWhileTickets, workSheet);
                    }

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

                    workSheet.Rows[107].Columns[columnWhileTickets].Value = numberSelledTickets * 20;
                    workSheet.Rows[107].Columns[columnWhileTickets].Style.Font.Height = 14;
                    workSheet.Rows[107].Columns[columnWhileTickets].Style.Font.Bold = true;
                    excelFunctions.BodersThinInt(107, columnWhileTickets, workSheet);
                    excelFunctions.CenterTextInt(107, columnWhileTickets, workSheet);
                    excelFunctions.PutBordersBlackInt(107, columnWhile, workSheet);

                    bool calculatedTicketsWeek = false;
                    while (!calculatedTicketsWeek)
                    {
                        string dayForCalculateCA = currentDay.DayOfWeek.ToString();
                        switch (dayForCalculateCA)
                        {
                            case "Monday":
                                
                                workSheet.Rows[94].Columns[columnWhileTickets].FormatString = BuiltinFormats.Percent2;
                                workSheet.Rows[94].Columns[columnWhileTickets].Value = ticketsVendus / ticketsDispos;
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
                                var cellAdress1TuesdaySum = workSheet.Rows[107].Columns[columnWhileTickets - 1].RangeAddressAsString;
                                var cellAdress2TuesdaySum = workSheet.Rows[107].Columns[columnWhileTickets].RangeAddressAsString;
                                var rangeTuesdayCA = workSheet[$"{cellAdress1TuesdaySum}:{cellAdress2TuesdaySum}"];
                                workSheet.Rows[109].Columns[columnWhileTickets - 1].Value = rangeTuesdayCA.Sum();
                                break;

                            case "Wednesday":
                                var cellAdress1WednesdaySum = workSheet.Rows[107].Columns[columnWhileTickets - 2].RangeAddressAsString;
                                var cellAdress2WednesdaySum = workSheet.Rows[107].Columns[columnWhileTickets].RangeAddressAsString;
                                var rangeWednesdayCA = workSheet[$"{cellAdress1WednesdaySum}:{cellAdress2WednesdaySum}"];
                                workSheet.Rows[109].Columns[columnWhileTickets - 2].Value = rangeWednesdayCA.Sum();
                                break;

                            case "Thursday":
                                var cellAdress1ThursdaySum = workSheet.Rows[107].Columns[columnWhileTickets - 3].RangeAddressAsString;
                                var cellAdress2ThursdaySum = workSheet.Rows[107].Columns[columnWhileTickets].RangeAddressAsString;
                                var rangeThursdayCA = workSheet[$"{cellAdress1ThursdaySum}:{cellAdress2ThursdaySum}"];
                                workSheet.Rows[109].Columns[columnWhileTickets - 3].Value = rangeThursdayCA.Sum();
                                break;

                            case "Friday":
                                var cellAdress1FridaySum = workSheet.Rows[107].Columns[columnWhileTickets - 4].RangeAddressAsString;
                                var cellAdress2FridaySum = workSheet.Rows[107].Columns[columnWhileTickets].RangeAddressAsString;
                                var rangeFridayCA = workSheet[$"{cellAdress1FridaySum}:{cellAdress2FridaySum}"];
                                workSheet.Rows[109].Columns[columnWhileTickets - 4].Value = rangeFridayCA.Sum();
                                break;

                            case "Saturday":
                                var cellAdress1SaturdaySum = workSheet.Rows[107].Columns[columnWhileTickets - 5].RangeAddressAsString;
                                var cellAdress2SaturdaySum = workSheet.Rows[107].Columns[columnWhileTickets].RangeAddressAsString;
                                var rangeSaturdayCA = workSheet[$"{cellAdress1SaturdaySum}:{cellAdress2SaturdaySum}"];
                                workSheet.Rows[109].Columns[columnWhileTickets - 5].Value = rangeSaturdayCA.Sum();
                                break;

                            case "Sunday":

                                var cellAdress1SundaySum = workSheet.Rows[107].Columns[columnWhileTickets - 6].RangeAddressAsString;
                                var cellAdress2SundaySum = workSheet.Rows[107].Columns[columnWhileTickets].RangeAddressAsString;
                                var rangeSundayCA = workSheet[$"{cellAdress1SundaySum}:{cellAdress2SundaySum}"];
                                workSheet.Rows[109].Columns[columnWhileTickets - 6].Value = rangeSundayCA.Sum();
                                break;
                        }
                        calculatedTicketsWeek = true;
                    }
                    calculatedTickets = true;
                }
                columnWhileTickets++;
            }
        }
    }
}
