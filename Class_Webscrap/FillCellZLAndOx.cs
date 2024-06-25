using ExcelLocalBiblioC;
using IronXL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Class_Webscrap
{
    public class FillCellZLAndOx
    {
        ExcelFunctions excelFunctions;

        public FillCellZLAndOx(WorkSheet sheet)
        {
            excelFunctions = new(sheet);
        }
        
        public void FillSelectedCellForZL(List<int> MaximumPlayer, List<int> NumberPlayer, WorkSheet sheet, int positionX, int[] positionY)
        {
            int incrementPlayer = 0;
            foreach (int PositionY in positionY)
            {
                sheet.Columns[positionX].Rows[PositionY].StringValue = $"{NumberPlayer[incrementPlayer]}/{MaximumPlayer[incrementPlayer]}";
                incrementPlayer++;
            }
        }

        public void ColorCell(List<int> MaximumPlayer, List<int> NumberPlayer, WorkSheet sheet, int positionX, int[] positionY)
        {
            foreach (int PositionY in positionY)
            {
                sheet.Columns[positionX].Rows[PositionY].Style.Font.Height = 14;
                string[] numberPlayerSplit = sheet.Columns[positionX].Rows[PositionY].StringValue.Split('/');
                decimal colornumber = Convert.ToDecimal(numberPlayerSplit[0]) / Convert.ToDecimal(numberPlayerSplit[1]);
                if (colornumber == 1m)
                {
                    sheet.Columns[positionX].Rows[PositionY].Style.BackgroundColor = "#e06666";

                    excelFunctions.CenterTextInt(PositionY, positionX, sheet);
                }
                if (colornumber < 0.5m)
                {
                    sheet.Columns[positionX].Rows[PositionY].Style.BackgroundColor = "#4a87e8";

                    excelFunctions.CenterTextInt(PositionY, positionX, sheet);
                }
                if (colornumber > 0.5m && colornumber < 1m)
                {
                    sheet.Columns[positionX].Rows[PositionY].Style.BackgroundColor = "#f6b26b";

                    excelFunctions.CenterTextInt(PositionY, positionX, sheet);
                }
                if (colornumber == 0.5m)
                {
                    sheet.Columns[positionX].Rows[PositionY].Style.BackgroundColor = "#fde6cd";

                    excelFunctions.CenterTextInt(PositionY, positionX, sheet);
                }

            }

            string ColumnStart = sheet.Columns[positionX].Rows[7].RangeAddressAsString;
            string ColumnEnd = sheet.Columns[positionX].Rows[91].RangeAddressAsString;

            if (sheet[$"{ColumnStart}:{ColumnEnd}"].IsEmpty)
            {
                
                sheet[$"{ColumnStart}:{ColumnEnd}"].Style.BackgroundColor = "#666666";
                
            }
        }
    }
}
