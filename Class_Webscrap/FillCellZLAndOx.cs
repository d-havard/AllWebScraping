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
        /// <summary>
        /// Take the informations of the maximum number of player and the actual number of player and make it in a string like this "00/00" 
        /// and put it in the excel
        /// </summary>
        /// <param name="MaximumPlayer"></param>
        /// <param name="NumberPlayer"></param>
        /// <param name="sheet"></param>
        /// <param name="positionX"></param>
        /// <param name="positionY"></param>
        public void FillSelectedCellForZL(List<int> MaximumPlayer, List<int> NumberPlayer, WorkSheet sheet, int positionX, int[] positionY)
        {
            int incrementPlayer = 0;
            foreach (int PositionY in positionY)
            {
                if (incrementPlayer < MaximumPlayer.Count)
                {
                    if (sheet.Columns[positionX].Rows[PositionY].Value != $"{NumberPlayer[incrementPlayer]}/{MaximumPlayer[incrementPlayer]}")
                    {
                        sheet.Columns[positionX].Rows[PositionY].StringValue = $"{NumberPlayer[incrementPlayer]}/{MaximumPlayer[incrementPlayer]}";
                        
                    }
                    incrementPlayer++;
                }
                
            }
        }
        /// <summary>
        /// Color the selected cell depending on how many player there are in one session
        /// </summary>
        /// <param name="MaximumPlayer"></param>
        /// <param name="NumberPlayer"></param>
        /// <param name="sheet"></param>
        /// <param name="positionX"></param>
        /// <param name="positionY"></param>
        public void ColorCell(List<int> MaximumPlayer, List<int> NumberPlayer, WorkSheet sheet, int positionX, int[] positionY)
        {
            foreach (int PositionY in positionY)
            {
                sheet.Columns[positionX].Rows[PositionY].Style.Font.Height = 14;
                var numberPlayer = sheet.Columns[positionX].Rows[PositionY];
                string[] numberPlayerSplit = numberPlayer.StringValue.Split('/');
                if (numberPlayerSplit[0] != "")
                {
                    decimal colornumber = Convert.ToDecimal(NumberPlayer[0]) / Convert.ToDecimal(numberPlayerSplit[1]);
                    if (colornumber == 1m)
                    {
                        if (sheet.Columns[positionX].Rows[PositionY].Style.BackgroundColor != "#e06666")
                        {
                            sheet.Columns[positionX].Rows[PositionY].Style.BackgroundColor = "#e06666";

                            excelFunctions.CenterTextInt(PositionY, positionX, sheet);
                        }

                    }
                    if (colornumber < 0.5m)
                    {
                        if (sheet.Columns[positionX].Rows[PositionY].Style.BackgroundColor != "#4a87e8")
                        {
                            sheet.Columns[positionX].Rows[PositionY].Style.BackgroundColor = "#4a87e8";

                            excelFunctions.CenterTextInt(PositionY, positionX, sheet);
                        }

                    }
                    if (colornumber > 0.5m && colornumber < 1m)
                    {
                        if (sheet.Columns[positionX].Rows[PositionY].Style.BackgroundColor != "#f6b26b")
                        {
                            sheet.Columns[positionX].Rows[PositionY].Style.BackgroundColor = "#f6b26b";

                            excelFunctions.CenterTextInt(PositionY, positionX, sheet);
                        }
                    }
                    if (colornumber == 0.5m)
                    {
                        if (sheet.Columns[positionX].Rows[PositionY].Style.BackgroundColor != "#fde6cd")
                        {
                            sheet.Columns[positionX].Rows[PositionY].Style.BackgroundColor = "#fde6cd";

                            excelFunctions.CenterTextInt(PositionY, positionX, sheet);
                        }

                    }
                }
                

            }

            //string ColumnStart = sheet.Columns[positionX].Rows[7].RangeAddressAsString;
            //string ColumnEnd = sheet.Columns[positionX].Rows[91].RangeAddressAsString;

            //if (sheet[$"{ColumnStart}:{ColumnEnd}"].IsEmpty)
            //{
            //    if (sheet[$"{ColumnStart}:{ColumnEnd}"].Style.BackgroundColor != "#666666")
            //    {
            //        sheet[$"{ColumnStart}:{ColumnEnd}"].Style.BackgroundColor = "#666666";
            //    }
            //}
        }
    }
}
