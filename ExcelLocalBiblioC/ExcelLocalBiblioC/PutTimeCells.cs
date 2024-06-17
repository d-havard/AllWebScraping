using IronXL.Formatting;
using IronXL.Styles;
using IronXL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLocalBiblioC
{
    public class PutTimeCells
    {
        WorkSheet workSheet;
        ExcelFunctions excelFunctions;

        public PutTimeCells(WorkSheet workSheet)
        {
            this.workSheet = workSheet;
            this.excelFunctions = new ExcelFunctions(workSheet);
        }

        public int PutActualDayInExcel(string today)
        {
            DateTime currentDay = DateTime.Today;
            ExcelFunctions excelFunctions = new ExcelFunctions(workSheet);
            bool dayPut = false;
            int columnWhile = 2;
            bool weekPut = false;
            while (!dayPut)
            {
                string[] cellBefore = workSheet.Rows[6].Columns[columnWhile - 1].Value.ToString().Split(' ');
                if (cellBefore[0] == today)
                {
                    dayPut = true;
                }
                else
                {
                    if (workSheet.Rows[6].Columns[columnWhile].IsEmpty && !dayPut)
                    {
                        string[] cellDate = workSheet.Rows[6].Columns[columnWhile - 1].Value.ToString().Split(' ');
                        if (workSheet.Rows[6].Columns[columnWhile + 1].IsEmpty && cellDate[0] != today)
                        {
                            int row = 6;
                            int column = columnWhile;
                            workSheet.Rows[6].Columns[columnWhile].StringValue = today;
                            workSheet.Rows[6].Columns[columnWhile].Style.BackgroundColor = "#b7b7b7";
                            excelFunctions.BordersThickInt(6, columnWhile, workSheet);
                            excelFunctions.CenterTextInt(6, columnWhile, workSheet);
                            workSheet.Rows[6].Columns[columnWhile].Style.Font.Bold = true;
                            workSheet.Rows[6].Columns[columnWhile].Style.Font.Height = 10;
                            workSheet.Columns[columnWhile].Width = 4000;
                            workSheet.Rows[6].Height = 300;

                            workSheet.Rows[92].Columns[columnWhile].StringValue = today;
                            workSheet.Rows[92].Columns[columnWhile].Style.BackgroundColor = "#b7b7b7";
                            excelFunctions.BordersThickInt(92, columnWhile, workSheet);
                            excelFunctions.CenterTextInt(92, columnWhile, workSheet);
                            workSheet.Rows[92].Columns[columnWhile].Style.Font.Bold = true;
                            workSheet.Rows[92].Columns[columnWhile].Style.Font.Height = 10;
                            workSheet.Columns[columnWhile].Width = 4000;
                            workSheet.Rows[92].Height = 300;

                            workSheet.Rows[106].Columns[columnWhile].StringValue = today;
                            workSheet.Rows[106].Columns[columnWhile].Style.BackgroundColor = "#b7b7b7";
                            excelFunctions.BordersThickInt(106, columnWhile, workSheet);
                            excelFunctions.CenterTextInt(106, columnWhile, workSheet);
                            workSheet.Rows[106].Columns[columnWhile].Style.Font.Bold = true;
                            workSheet.Rows[106].Columns[columnWhile].Style.Font.Height = 10;
                            workSheet.Columns[columnWhile].Width = 4000;
                            workSheet.Rows[106].Height = 300;



                            string dayForMerge = currentDay.DayOfWeek.ToString();
                            switch (dayForMerge)
                            {
                                case "Monday":
                                    while (!weekPut)
                                    {
                                        if (workSheet.Rows[94].Columns[columnWhile].IsEmpty)
                                        {
                                            var cellAdress1Merge = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2Merge = workSheet.Rows[94].Columns[columnWhile + 6].RangeAddressAsString;
                                            var mergeWeek = workSheet.Merge($"{cellAdress1Merge}:{cellAdress2Merge}");
                                            var cellAdress1Avg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2Avg = workSheet.Rows[93].Columns[columnWhile + 6].RangeAddressAsString;
                                            var rangeAvg = workSheet[$"{cellAdress1Avg}:{cellAdress2Avg}"];
                                            var cellAdressString = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressString}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressString}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressString, workSheet);

                                            var cellAdress1MergeCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2MergeCA = workSheet.Rows[109].Columns[columnWhile + 6].RangeAddressAsString;
                                            var mergeWeekCA = workSheet.Merge($"{cellAdress1MergeCA}:{cellAdress2MergeCA}");
                                            var cellAdress1AvgCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2AvgCA = workSheet.Rows[109].Columns[columnWhile + 6].RangeAddressAsString;
                                            var rangeAvgCA = workSheet[$"{cellAdress1AvgCA}:{cellAdress2AvgCA}"];
                                            var cellAdressStringCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressStringCA}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressStringCA}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressStringCA, workSheet);
                                            if (workSheet[$"{cellAdressString}"].IsEmpty)
                                            {

                                            }
                                            else
                                            {
                                                workSheet[$"{cellAdressString}"].Value = rangeAvg.Avg();
                                            }
                                            weekPut = true;
                                        }
                                        else
                                        {
                                            var cellAdress1Avg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2Avg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                            var rangeAvg = workSheet[$"{cellAdress1Avg}:{cellAdress2Avg}"];
                                            var cellAdressString = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressString}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressString}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressString, workSheet);
                                            workSheet[$"{cellAdressString}"].Value = rangeAvg.Avg();
                                            weekPut = true;
                                        }
                                    }
                                    break;

                                case "Tuesday":
                                    while (!weekPut)
                                    {
                                        if (workSheet.Rows[94].Columns[columnWhile].IsEmpty)
                                        {
                                            var cellAdress1Merge = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2Merge = workSheet.Rows[94].Columns[columnWhile + 5].RangeAddressAsString;
                                            var mergeWeek = workSheet.Merge($"{cellAdress1Merge}:{cellAdress2Merge}");
                                            var cellAdress1Avg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2Avg = workSheet.Rows[93].Columns[columnWhile + 5].RangeAddressAsString;
                                            var rangeAvg = workSheet[$"{cellAdress1Avg}:{cellAdress2Avg}"];
                                            var cellAdressString = workSheet.Rows[94].Columns[2].RangeAddressAsString;
                                            workSheet[$"{cellAdressString}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressString}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressString, workSheet);

                                            var cellAdress1MergeCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2MergeCA = workSheet.Rows[109].Columns[columnWhile + 5].RangeAddressAsString;
                                            var mergeWeekCA = workSheet.Merge($"{cellAdress1MergeCA}:{cellAdress2MergeCA}");
                                            var cellAdress1AvgCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2AvgCA = workSheet.Rows[109].Columns[columnWhile + 5].RangeAddressAsString;
                                            var rangeAvgCA = workSheet[$"{cellAdress1AvgCA}:{cellAdress2AvgCA}"];
                                            var cellAdressStringCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressStringCA}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressStringCA}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressStringCA, workSheet);
                                            if (workSheet[$"{cellAdressString}"].IsEmpty)
                                            {

                                            }
                                            else
                                            {
                                                workSheet[$"{cellAdressString}"].Value = rangeAvg.Avg();
                                            }
                                            weekPut = true;
                                        }
                                        else
                                        {
                                            var cellAdress1Avg = workSheet.Rows[93].Columns[columnWhile - 1].RangeAddressAsString;
                                            var cellAdress2Avg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                            var rangeAvg = workSheet[$"{cellAdress1Avg}:{cellAdress2Avg}"];
                                            var cellAdressString = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressString}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressString}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressString, workSheet);
                                            workSheet[$"{cellAdressString}"].Value = rangeAvg.Avg();
                                            weekPut = true;
                                        }
                                    }
                                    break;

                                case "Wednesday":
                                    while (!weekPut)
                                    {
                                        if (workSheet.Rows[94].Columns[columnWhile].IsEmpty)
                                        {
                                            var cellAdress1Merge = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2Merge = workSheet.Rows[94].Columns[columnWhile + 4].RangeAddressAsString;
                                            var mergeWeek = workSheet.Merge($"{cellAdress1Merge}:{cellAdress2Merge}");
                                            var cellAdress1Avg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2Avg = workSheet.Rows[93].Columns[columnWhile + 4].RangeAddressAsString;
                                            var rangeAvg = workSheet[$"{cellAdress1Avg}:{cellAdress2Avg}"];
                                            var cellAdressString = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressString}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressString}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressString, workSheet);

                                            var cellAdress1MergeCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2MergeCA = workSheet.Rows[109].Columns[columnWhile + 4].RangeAddressAsString;
                                            var mergeWeekCA = workSheet.Merge($"{cellAdress1MergeCA}:{cellAdress2MergeCA}");
                                            var cellAdress1AvgCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2AvgCA = workSheet.Rows[109].Columns[columnWhile + 4].RangeAddressAsString;
                                            var rangeAvgCA = workSheet[$"{cellAdress1AvgCA}:{cellAdress2AvgCA}"];
                                            var cellAdressStringCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressStringCA}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressStringCA}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressStringCA, workSheet);
                                            if (workSheet[$"{cellAdressString}"].IsEmpty)
                                            {

                                            }
                                            else
                                            {
                                                workSheet[$"{cellAdressString}"].Value = rangeAvg.Avg();
                                            }
                                            weekPut = true;
                                        }
                                        else
                                        {
                                            var cellAdress1Avg = workSheet.Rows[93].Columns[columnWhile - 2].RangeAddressAsString;
                                            var cellAdress2Avg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                            var rangeAvg = workSheet[$"{cellAdress1Avg}:{cellAdress2Avg}"];
                                            var cellAdressString = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressString}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressString}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressString, workSheet);
                                            workSheet[$"{cellAdressString}"].Value = rangeAvg.Avg();
                                            weekPut = true;
                                        }
                                    }
                                    break;

                                case "Thursday":
                                    while (!weekPut)
                                    {
                                        if (workSheet.Rows[94].Columns[columnWhile].IsEmpty)
                                        {
                                            var cellAdress1Merge = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2Merge = workSheet.Rows[94].Columns[columnWhile + 3].RangeAddressAsString;
                                            var mergeWeek = workSheet.Merge($"{cellAdress1Merge}:{cellAdress2Merge}");
                                            var cellAdress1Avg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2Avg = workSheet.Rows[93].Columns[columnWhile + 3].RangeAddressAsString;
                                            var rangeAvg = workSheet[$"{cellAdress1Avg}:{cellAdress2Avg}"];
                                            var cellAdressString = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressString}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressString}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressString, workSheet);

                                            var cellAdress1MergeCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2MergeCA = workSheet.Rows[109].Columns[columnWhile + 3].RangeAddressAsString;
                                            var mergeWeekCA = workSheet.Merge($"{cellAdress1MergeCA}:{cellAdress2MergeCA}");
                                            var cellAdress1AvgCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2AvgCA = workSheet.Rows[109].Columns[columnWhile + 3].RangeAddressAsString;
                                            var rangeAvgCA = workSheet[$"{cellAdress1AvgCA}:{cellAdress2AvgCA}"];
                                            var cellAdressStringCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressStringCA}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressStringCA}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressStringCA, workSheet);
                                            if (workSheet[$"{cellAdressString}"].IsEmpty)
                                            {

                                            }
                                            else
                                            {
                                                workSheet[$"{cellAdressString}"].Value = rangeAvg.Avg();
                                            }
                                            weekPut = true;
                                        }
                                        else
                                        {
                                            var cellAdress1Avg = workSheet.Rows[93].Columns[columnWhile - 3].RangeAddressAsString;
                                            var cellAdress2Avg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                            var rangeAvg = workSheet[$"{cellAdress1Avg}:{cellAdress2Avg}"];
                                            var cellAdressString = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressString}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressString}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressString, workSheet);
                                            workSheet[$"{cellAdressString}"].Value = rangeAvg.Avg();
                                            weekPut = true;
                                        }
                                    }
                                    break;

                                case "Friday":
                                    while (!weekPut)
                                    {
                                        if (workSheet.Rows[94].Columns[columnWhile].IsEmpty)
                                        {
                                            var cellAdress1Merge = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2Merge = workSheet.Rows[94].Columns[columnWhile + 2].RangeAddressAsString;
                                            var mergeWeek = workSheet.Merge($"{cellAdress1Merge}:{cellAdress2Merge}");
                                            var cellAdress1Avg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2Avg = workSheet.Rows[93].Columns[columnWhile + 2].RangeAddressAsString;
                                            var rangeAvg = workSheet[$"{cellAdress1Avg}:{cellAdress2Avg}"];
                                            var cellAdressString = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressString}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressString}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressString, workSheet);

                                            var cellAdress1MergeCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2MergeCA = workSheet.Rows[109].Columns[columnWhile + 2].RangeAddressAsString;
                                            var mergeWeekCA = workSheet.Merge($"{cellAdress1MergeCA}:{cellAdress2MergeCA}");
                                            var cellAdress1AvgCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2AvgCA = workSheet.Rows[109].Columns[columnWhile + 2].RangeAddressAsString;
                                            var rangeAvgCA = workSheet[$"{cellAdress1AvgCA}:{cellAdress2AvgCA}"];
                                            var cellAdressStringCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressStringCA}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressStringCA}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressStringCA, workSheet);
                                            if (workSheet[$"{cellAdressString}"].IsEmpty)
                                            {

                                            }
                                            else
                                            {
                                                workSheet[$"{cellAdressString}"].Value = rangeAvg.Avg();
                                            }
                                            weekPut = true;
                                        }
                                        else
                                        {
                                            var cellAdress1Avg = workSheet.Rows[93].Columns[columnWhile - 4].RangeAddressAsString;
                                            var cellAdress2Avg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                            var rangeAvg = workSheet[$"{cellAdress1Avg}:{cellAdress2Avg}"];
                                            var cellAdressString = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressString}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressString}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressString, workSheet);
                                            workSheet[$"{cellAdressString}"].Value = rangeAvg.Avg();
                                            weekPut = true;
                                        }
                                    }
                                    break;

                                case "Saturday":
                                    while (!weekPut)
                                    {
                                        if (workSheet.Rows[94].Columns[columnWhile].IsEmpty)
                                        {
                                            var cellAdress1Merge = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2Merge = workSheet.Rows[94].Columns[columnWhile + 1].RangeAddressAsString;
                                            var mergeWeek = workSheet.Merge($"{cellAdress1Merge}:{cellAdress2Merge}");
                                            var cellAdress1Avg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2Avg = workSheet.Rows[93].Columns[columnWhile + 1].RangeAddressAsString;
                                            var rangeAvg = workSheet[$"{cellAdress1Avg}:{cellAdress2Avg}"];
                                            var cellAdressString = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressString}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressString}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressString, workSheet);

                                            var cellAdress1MergeCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2MergeCA = workSheet.Rows[109].Columns[columnWhile + 1].RangeAddressAsString;
                                            var mergeWeekCA = workSheet.Merge($"{cellAdress1MergeCA}:{cellAdress2MergeCA}");
                                            var cellAdress1AvgCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2AvgCA = workSheet.Rows[109].Columns[columnWhile + 1].RangeAddressAsString;
                                            var rangeAvgCA = workSheet[$"{cellAdress1AvgCA}:{cellAdress2AvgCA}"];
                                            var cellAdressStringCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressStringCA}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressStringCA}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressStringCA, workSheet);
                                            if (workSheet[$"{cellAdressString}"].IsEmpty)
                                            {

                                            }
                                            else
                                            {
                                                workSheet[$"{cellAdressString}"].Value = rangeAvg.Avg();
                                            }
                                            weekPut = true;
                                        }
                                        else
                                        {
                                            var cellAdress1Avg = workSheet.Rows[93].Columns[columnWhile - 5].RangeAddressAsString;
                                            var cellAdress2Avg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                            var rangeAvg = workSheet[$"{cellAdress1Avg}:{cellAdress2Avg}"];
                                            var cellAdressString = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressString}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressString}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressString, workSheet);
                                            workSheet[$"{cellAdressString}"].Value = rangeAvg.Avg();
                                            weekPut = true;
                                        }
                                    }
                                    break;

                                case "Sunday":
                                    while (!weekPut)
                                    {
                                        if (workSheet.Rows[94].Columns[columnWhile].IsEmpty)
                                        {
                                            var cellAdress1Merge = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2Merge = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress1Avg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2Avg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                            var rangeAvg = workSheet[$"{cellAdress1Avg}:{cellAdress2Avg}"];
                                            var cellAdressString = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressString}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressString}"].Style.Font.Bold = true;
                                            workSheet[$"{cellAdressString}"].Style.VerticalAlignment = VerticalAlignment.Center;
                                            workSheet[$"{cellAdressString}"].Style.HorizontalAlignment = HorizontalAlignment.Center;

                                            var cellAdress1MergeCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2MergeCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress1AvgCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            var cellAdress2AvgCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            var rangeAvgCA = workSheet[$"{cellAdress1AvgCA}:{cellAdress2AvgCA}"];
                                            var cellAdressStringCA = workSheet.Rows[109].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressStringCA}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressStringCA}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressStringCA, workSheet);
                                            if (workSheet[$"{cellAdressString}"].IsEmpty)
                                            {
                                                workSheet[$"{cellAdressString}"].Value = rangeAvg.Avg();
                                            }
                                            else
                                            {
                                                workSheet[$"{cellAdressString}"].Value = rangeAvg.Avg();
                                            }
                                            weekPut = true;
                                        }
                                        else
                                        {
                                            var cellAdress1Avg = workSheet.Rows[93].Columns[columnWhile - 6].RangeAddressAsString;
                                            var cellAdress2Avg = workSheet.Rows[93].Columns[columnWhile].RangeAddressAsString;
                                            var rangeAvg = workSheet[$"{cellAdress1Avg}:{cellAdress2Avg}"];
                                            var cellAdressString = workSheet.Rows[94].Columns[columnWhile].RangeAddressAsString;
                                            workSheet[$"{cellAdressString}"].Style.Font.Height = 18;
                                            workSheet[$"{cellAdressString}"].Style.Font.Bold = true;
                                            excelFunctions.CenterTextString(cellAdressString, workSheet);
                                            workSheet[$"{cellAdressString}"].Value = rangeAvg.Avg();
                                            weekPut = true;
                                        }
                                    }
                                    break;

                            }
                            dayPut = true;
                        }
                    }
                }
                columnWhile++;
            }
            return columnWhile;
        }

        
    }
}