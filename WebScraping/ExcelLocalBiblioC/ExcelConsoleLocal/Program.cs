using ExcelLocalBiblioC;

namespace ExcelConsoleLocal
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string path = "C:\\Stage\\AllWebScraping_VirtualGame\\dataFormatts";
            ExcelStructure excelStructure = new ExcelStructure(path);
            DataExcel dataExcel = new DataExcel();

            var workSheet = excelStructure.VerificationIfFileExist(path);
            //excelStructure.CreateStructure(path);
            
            dataExcel.PutDataExcel(workSheet);
            excelStructure.SaveFile();
        }
    }
}
