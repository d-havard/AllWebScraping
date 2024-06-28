using IronXL;
using IronXL.Formatting;
using IronXL.Styles;
using NUnit.Framework.Internal.Execution;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Class_Webscrap
{
    public class FillCell
    {
        /// <summary>
        /// Take the informations of the maximum number of player and the actual number of player and make it in a string like this "00/00" 
        /// and put it in the excel
        /// </summary>
        /// <param name="MaximumPlayer"></param>
        /// <param name="NumberPlayer"></param>
        /// <param name="sheet"></param>
        /// <param name="positionX"></param>
        /// <param name="positionY"></param>
        public void FillSelectedCell(List<int> MaximumPlayer, List<int> NumberPlayer, WorkSheet sheet, int positionX, int[] positionY)
        {
            int incrementPlayer = 0;
            foreach (int PositionY in positionY)
            {
                if (MaximumPlayer[incrementPlayer] != 0)
                {
                    sheet.Columns[positionX].Rows[PositionY].StringValue = $"{NumberPlayer[incrementPlayer]}/{MaximumPlayer[incrementPlayer]}";
                }
                else
                {
                    sheet.Columns[positionX].Rows[PositionY].StringValue = "Priva";
                }
                      
                incrementPlayer++;
            }
        }


        public void colorCell(List<int> MaximumPlayer, List<int> NumberPlayer, WorkSheet sheet, int positionX, int[] positionY, List<bool> battlepassPlayers, List<bool> peakHours)
        {
            int incrementPlayer = 0;
            foreach (int PositionY in positionY)
            {
                if (!peakHours[incrementPlayer])
                {
                    sheet.Columns[positionX].Rows[PositionY].Style.TopBorder.Type = BorderType.Thick;
                    sheet.Columns[positionX].Rows[PositionY].Style.RightBorder.Type = BorderType.Thick;
                    sheet.Columns[positionX].Rows[PositionY].Style.LeftBorder.Type = BorderType.Thick;
                    sheet.Columns[positionX].Rows[PositionY].Style.BottomBorder.Type = BorderType.Thick;
                    sheet.Columns[positionX].Rows[PositionY].Style.BottomBorder.Color = "#00ff02";
                    sheet.Columns[positionX].Rows[PositionY].Style.LeftBorder.Color = "#00ff02";
                    sheet.Columns[positionX].Rows[PositionY].Style.RightBorder.Color = "#00ff02";
                    sheet.Columns[positionX].Rows[PositionY].Style.TopBorder.Color = "#00ff02";

                }
                if (sheet.Columns[positionX].Rows[PositionY].StringValue == "Priva")
                {
                    sheet.Columns[positionX].Rows[PositionY].Style.Font.Bold = true;
                    sheet.Columns[positionX].Rows[PositionY].Style.Font.Height = 18;
                    sheet.Columns[positionX].Rows[PositionY].Style.VerticalAlignment = IronXL.Styles.VerticalAlignment.Center;
                    sheet.Columns[positionX].Rows[PositionY].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center;

                }
                else
                {
                    string[] numberPlayerSplit = sheet.Columns[positionX].Rows[PositionY].StringValue.Split('/');
                    decimal colornumber = Convert.ToDecimal(numberPlayerSplit[0]) / Convert.ToDecimal(numberPlayerSplit[1]);
                    if (battlepassPlayers[incrementPlayer])
                    {
                        sheet.Columns[positionX].Rows[PositionY].Style.BackgroundColor = "#9900ff";
                        sheet.Columns[positionX].Rows[PositionY].Style.Font.Height = 14;
                        sheet.Columns[positionX].Rows[PositionY].Style.VerticalAlignment = IronXL.Styles.VerticalAlignment.Center;
                        sheet.Columns[positionX].Rows[PositionY].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center;
                        
                    }

                    else
                    {
                        if (colornumber == 1m)
                        {
                            sheet.Columns[positionX].Rows[PositionY].Style.BackgroundColor = "#e06666";
                            sheet.Columns[positionX].Rows[PositionY].Style.Font.Height = 14;
                            sheet.Columns[positionX].Rows[PositionY].Style.VerticalAlignment = IronXL.Styles.VerticalAlignment.Center;
                            sheet.Columns[positionX].Rows[PositionY].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center;
                           
                        }
                        if (colornumber < 0.5m)
                        {
                            sheet.Columns[positionX].Rows[PositionY].Style.BackgroundColor = "#4a87e8";
                            sheet.Columns[positionX].Rows[PositionY].Style.Font.Height = 14;
                            sheet.Columns[positionX].Rows[PositionY].Style.VerticalAlignment = IronXL.Styles.VerticalAlignment.Center;
                            sheet.Columns[positionX].Rows[PositionY].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center;
                            
                        }
                        if (colornumber > 0.5m && colornumber < 1m)
                        {
                            sheet.Columns[positionX].Rows[PositionY].Style.BackgroundColor = "#f6b26b";
                            sheet.Columns[positionX].Rows[PositionY].Style.Font.Height = 14;
                            sheet.Columns[positionX].Rows[PositionY].Style.VerticalAlignment = IronXL.Styles.VerticalAlignment.Center;
                            sheet.Columns[positionX].Rows[PositionY].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center;
                            
                        }
                        if (colornumber == 0.5m)
                        {
                            sheet.Columns[positionX].Rows[PositionY].Style.BackgroundColor = "#fde6cd";
                            sheet.Columns[positionX].Rows[PositionY].Style.Font.Height = 14;
                            sheet.Columns[positionX].Rows[PositionY].Style.VerticalAlignment = IronXL.Styles.VerticalAlignment.Center;
                            sheet.Columns[positionX].Rows[PositionY].Style.HorizontalAlignment = IronXL.Styles.HorizontalAlignment.Center;
                            
                        }
                    }
                    
                }
                
                incrementPlayer++;
                
            }
            
            
        }
    }
}
