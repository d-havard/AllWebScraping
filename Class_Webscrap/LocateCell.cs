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

        public int[] LocateCellYPosition(List<string> StartedHours, WorkSheet sheet)
        {
            int locY = 7;
           
            List<int> positionYList = new();
            var range = sheet["B8:B92"];
            foreach (string startedhour in StartedHours)
            {
                string starthour = startedhour + ":00";
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
