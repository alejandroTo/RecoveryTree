using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RecoveryTree
{
    public static class XlHelper
    {
        
        public static Dictionary<string, int> GetXlHeadersTableHorizontal(Worksheet XlSheet)
        {
            Dictionary<string, int> Headers = new Dictionary<string, int>();
            int xlCol = 1;
            try
            {
                while (!(XlSheet.Cells[xlCol,1].Value == null))
                {
                    Headers.Add(Convert.ToString(XlSheet.Cells[xlCol,1].value), xlCol);
                    xlCol++;
                }
                return Headers;
            }
            catch (Exception ex)
            {
                return Headers;
            }

        }
        public static Dictionary<string, int> GetXlHeadersTableVertical(Worksheet XlSheet)
        {
            Dictionary<string, int> Headers = new Dictionary<string, int>();
            int xlCol = 1;
            try
            {
                while (!(XlSheet.Cells[1, xlCol].Value == null))
                {
                    Headers.Add(Convert.ToString(XlSheet.Cells[1, xlCol].Value), xlCol);
                    xlCol++;
                    //1,i+1
                }
                return Headers;
            }
            catch (Exception ex)
            {
                return Headers;
            }

        }
    }
}
