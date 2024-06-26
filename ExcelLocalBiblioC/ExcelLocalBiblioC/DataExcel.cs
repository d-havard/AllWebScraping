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
    public class DataExcel
    {
        private string OriginalFilePath;
        WorkSheet workSheet;
        WorkBook workBook;
        public DataExcel() 
        {
            
        }

        public int PutDataExcel(WorkSheet workSheet) 
        {
            ExcelFunctions excelFunctions = new ExcelFunctions(workSheet);
            PutTimeCells timeFunctions = new PutTimeCells(workSheet);
            ExcelCalculs excelCalculs = new ExcelCalculs(workSheet);
            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;
            int day = DateTime.Now.Day;
            string zeroBeforeTenDay = "0";
            string zeroBeforeTenMonth = "0";
            string today = $"{zeroBeforeTenDay}{day}/{zeroBeforeTenMonth}{month}/{year}";

            if (day < 10 && month < 10)
            {
                today = $"{zeroBeforeTenMonth}{day}/{zeroBeforeTenMonth}{month}/{year}";
            }
            if (day < 10 && month >= 10)
            {
                today = $"{zeroBeforeTenDay}{day}/{month}/{year}";
            }
            if (day >= 10 && month < 10)
            {
                today = $"{day}/{zeroBeforeTenMonth}{month}/{year}";
            }
            if (day >= 10 && month >= 10)
            {
                today = $"{day}/{month}/{year}";
            }

            int columnWhile = timeFunctions.PutActualDayInExcel(today, workSheet);
            excelCalculs.CalculateTickets(columnWhile);
            return columnWhile;
        }
    }
}