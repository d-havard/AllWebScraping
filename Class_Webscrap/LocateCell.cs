using IronXL;
using IronXL.Drawing;
using OpenQA.Selenium.DevTools.V123.Page;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Class_Webscrap
{
    public class LocateCell
    {

        /// <summary>
        /// Locate the cell where is locate the date of today
        /// </summary>
        /// <param name="date"></param>
        /// <param name="sheet"></param>
        /// <returns>The position of today in the excel</returns>
        public int locateCellXPosition(string date, WorkSheet sheet)
        {
            int positionxtrouver = 0;
            int positionx = 2;
            //var range = sheet["C7:Q7"];
            
            while (positionxtrouver == 0)
            {
                string[] splitDate = sheet.Rows[6].Columns[positionx].Value.ToString().Split(' ');
                if (date != splitDate[0])
                {
                    positionx++;
                }
                else
                {
                    positionxtrouver = positionx;
                }
                
            }
            return positionxtrouver;
        }

        /// <summary>
        /// Find the positions of all the hours found in the JSON, and locate them in the excel
        /// </summary>
        /// <param name="StartedHours"></param>
        /// <param name="sheet"></param>
        /// <returns>Positions of all the started hours of each sessions</returns>
        public int[] LocateCellYPosition(List<string> StartedHours, WorkSheet sheet, string nomSheet)
        {
           
            List<int> positionYList = new();
            var range = sheet["B8:B92"];
            foreach (string startedhour in StartedHours)
            {
                
                string[] startedhourSplit = startedhour.Split(':');
                string starthour = startedhour;
                if (nomSheet.Contains("OXMOZ"))
                {
                    string[] hoursplit = startedhour.Split(" ");
                    starthour = hoursplit[1];
                }
                if (startedhour.Length == 5)
                {
                    starthour = startedhour + ":00";
                }
                if (startedhourSplit[1].Remove(0, 1) == "5")
                {
                    int stringconverted = Int32.Parse(startedhourSplit[1]);
                    stringconverted += 5;
                    starthour = startedhourSplit[0] + stringconverted.ToString() + startedhourSplit[2];
                }
                int locY = 7;
                foreach (var cell in range)
                {
                    string[] split = cell.Value.ToString().Split(' ');
                    
                        if (starthour == split[1])
                        {
                            positionYList.Add(locY);
                            locY = 7;
                            break;
                        }
                        locY++;
                }
                
            }
            int [] positionY = positionYList.ToArray();
            return positionY;
        }

        
    }
}
