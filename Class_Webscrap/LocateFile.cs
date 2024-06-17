using System.IO;
using IronXL;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Class_Webscrap
{
    public class LocateFile
    {
        private string OriginalFilePath;
        private string SaveFilePath;
        WorkSheet sheet;
        WorkBook workbook;

        /// <summary>
        /// The constructor for the webscraping, were you put the url of the website,
        /// the location of the original file and the location were you want to save your file 
        /// (where you have to put a @ in front of the two file location and put double \ in the two path)
        /// </summary>
        /// <param name="OriginalFilePath"></param>
        /// <param name="SaveFilePath"></param>
        public LocateFile(string OriginalFilePath,string SaveFilePath)
        {
            this.OriginalFilePath = OriginalFilePath;
            this.SaveFilePath = SaveFilePath;
            
        }
        

        /// <summary>
        /// Verify if the file exist in the path given. 
        /// Create an excel file with the worksheet name given if the file doesn't exist in the file,
        /// Or load the escel file if it exist.
        /// </summary>
        /// <param name="OriginalFilePath"></param>
        /// <param name="Worksheet"></param>
        public WorkSheet VerificationIfFileExist(string Worksheet)
        {
            

            if (!File.Exists(OriginalFilePath))
            {
                workbook = WorkBook.Create(ExcelFileFormat.XLSX);
                sheet = workbook.CreateWorkSheet(Worksheet);
            }
            else
            {
                workbook = WorkBook.LoadExcel(OriginalFilePath);
                sheet = workbook.GetWorkSheet(Worksheet);
            }

            return sheet;
        }

        /// <summary>
        /// Save the worksheet in the file.
        /// </summary>
        /// <param name="worksheet"></param>
        public void SaveFile(WorkSheet worksheet)
        {
            workbook.SaveAs(SaveFilePath);
        }
    }
}
