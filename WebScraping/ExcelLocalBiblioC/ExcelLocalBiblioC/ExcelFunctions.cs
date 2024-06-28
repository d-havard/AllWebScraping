using IronXL.Formatting;
using IronXL.Styles;
using IronXL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLocalBiblioC
{
    public class ExcelFunctions
    {
        WorkSheet workSheet;
        public ExcelFunctions(WorkSheet workSheet) 
        {
            this.workSheet = workSheet;
        }
        
        /// <summary>
        /// The function make all the 4 borders of the cell thick. The cell is gived by a string.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="workSheet"></param>
        public void BordersThickString(string cell, WorkSheet workSheet)
        {
            workSheet[$"{cell}"].Style.TopBorder.Type = BorderType.Thick;
            workSheet[$"{cell}"].Style.RightBorder.Type = BorderType.Thick;
            workSheet[$"{cell}"].Style.LeftBorder.Type = BorderType.Thick;
            workSheet[$"{cell}"].Style.BottomBorder.Type = BorderType.Thick;
        }

        /// <summary>
        /// Same that BordersThickString but we giving the row and the column in int. ! The number of the row and column beging at 0 !
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="workSheet"></param>
        public void BordersThickInt(int row, int column, WorkSheet workSheet)
        {
            workSheet.Rows[row].Columns[column].Style.TopBorder.Type = BorderType.Thick;
            workSheet.Rows[row].Columns[column].Style.RightBorder.Type = BorderType.Thick;
            workSheet.Rows[row].Columns[column].Style.LeftBorder.Type = BorderType.Thick;
            workSheet.Rows[row].Columns[column].Style.BottomBorder.Type = BorderType.Thick;
        }

        public void BodersThinInt(int row, int column, WorkSheet workSheet)
        {
            workSheet.Rows[row].Columns[column].Style.TopBorder.Type = BorderType.Thin;
            workSheet.Rows[row].Columns[column].Style.RightBorder.Type = BorderType.Thin;
            workSheet.Rows[row].Columns[column].Style.LeftBorder.Type = BorderType.Thin;
            workSheet.Rows[row].Columns[column].Style.BottomBorder.Type = BorderType.Thin;
        }

        public void BordersThinString(string cell, WorkSheet workSheet)
        {
            workSheet[$"{cell}"].Style.TopBorder.Type = BorderType.Thin;
            workSheet[$"{cell}"].Style.RightBorder.Type = BorderType.Thin;
            workSheet[$"{cell}"].Style.LeftBorder.Type = BorderType.Thin;
            workSheet[$"{cell}"].Style.BottomBorder.Type = BorderType.Thin;
        }
            
        /// <summary>
        /// Center the text of the cell of the workSheet we want, cell put with a string 
        /// (ex: "B17", Sheet1)
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="workSheet"></param>
        public void CenterTextString(string cell, WorkSheet workSheet)
        {
            workSheet[$"{cell}"].Style.VerticalAlignment = VerticalAlignment.Center;
            workSheet[$"{cell}"].Style.HorizontalAlignment = HorizontalAlignment.Center;
        }

        /// <summary>
        /// Center the text of the cell choose by giving the column and the row of the cell in int. ! The number of the row and column beging at 0 !
        /// (ex : row = 6 & column = 10)
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="workSheet"></param>
        public void CenterTextInt(int row, int column, WorkSheet workSheet)
        {
            workSheet.Rows[row].Columns[column].Style.VerticalAlignment = VerticalAlignment.Center;
            workSheet.Rows[row].Columns[column].Style.HorizontalAlignment = HorizontalAlignment.Center;
        }

        /// <summary>
        /// Put a green thick all around the cell we want in the WorkSheet we want, cell put with a string.
        /// (ex: "B17", Sheet1)
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="workSheet"></param>
        public void GreenThickBordersString(string cell, WorkSheet workSheet)
        {
            workSheet[$"{cell}"].Style.TopBorder.Type = BorderType.Thick;
            workSheet[$"{cell}"].Style.RightBorder.Type = BorderType.Thick;
            workSheet[$"{cell}"].Style.LeftBorder.Type = BorderType.Thick;
            workSheet[$"{cell}"].Style.BottomBorder.Type = BorderType.Thick;
            workSheet[$"{cell}"].Style.BottomBorder.Color = "#00ff02";
            workSheet[$"{cell}"].Style.LeftBorder.Color = "#00ff02";
            workSheet[$"{cell}"].Style.RightBorder.Color = "#00ff02";
            workSheet[$"{cell}"].Style.TopBorder.Color = "#00ff02";
        }

        /// <summary>
        /// Same that GreenThickBorders but we giving the row and the column with a int. ! The number of the row and column beging at 0 !
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="workSheet"></param>
        public void GreenThickBordersInt(int row, int column, WorkSheet workSheet)
        {
            workSheet.Rows[row].Columns[column].Style.TopBorder.Type = BorderType.Thick;
            workSheet.Rows[row].Columns[column].Style.RightBorder.Type = BorderType.Thick;
            workSheet.Rows[row].Columns[column].Style.LeftBorder.Type = BorderType.Thick;
            workSheet.Rows[row].Columns[column].Style.BottomBorder.Type = BorderType.Thick;
            workSheet.Rows[row].Columns[column].Style.BottomBorder.Color = "#00ff02";
            workSheet.Rows[row].Columns[column].Style.LeftBorder.Color = "#00ff02";
            workSheet.Rows[row].Columns[column].Style.RightBorder.Color = "#00ff02";
            workSheet.Rows[row].Columns[column].Style.TopBorder.Color = "#00ff02";
        }
        
        public void PutBordersBlackString(string cell, WorkSheet workSheet)
        {
            workSheet[$"{cell}"].Style.BottomBorder.Color = "#000000";
            workSheet[$"{cell}"].Style.LeftBorder.Color = "#000000";
            workSheet[$"{cell}"].Style.RightBorder.Color = "#000000";
            workSheet[$"{cell}"].Style.TopBorder.Color = "#000000";
        }
        public void PutBordersBlackInt(int row, int column, WorkSheet workSheet)
        {
            workSheet.Rows[row].Columns[column].Style.BottomBorder.Color = "#000000";
            workSheet.Rows[row].Columns[column].Style.LeftBorder.Color = "#000000";
            workSheet.Rows[row].Columns[column].Style.RightBorder.Color = "#000000";
            workSheet.Rows[row].Columns[column].Style.TopBorder.Color = "#000000";
        }
    }
}
