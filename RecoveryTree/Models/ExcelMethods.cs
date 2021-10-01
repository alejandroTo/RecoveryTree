using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace RecoveryTree.Models
{
    class ExcelMethods
    {

        public void FillExcel(int i, int j, string text)
        {
            Range Celda = Globals.wShe.Cells[i, j];
            Celda.Value = text.Trim() == "" ? " " : text.Trim();
        }
        public void OpenExcel()
        {
            try
            {
                Globals.xlApp = new Microsoft.Office.Interop.Excel.Application();
                Globals.xlWk = Globals.xlApp.Workbooks.Add();
                Globals.wShe = (Worksheet)Globals.xlWk.Sheets.Item[1];
                Globals.xlApp.Visible = true;
                Globals.xlApp.DisplayAlerts = true;

            }
            catch (System.Exception ex)
            {

            }

        }

        public int GetHorizontalValue(Dictionary<string, int> headersHorizantal, string value)
        {
            foreach (var item in headersHorizantal)
            {
                if (item.Key.Contains(value))
                {
                    return item.Value;
                }
            }
            return 0;
        }
        public void GetHeaders()
        {
            Dictionary<string, int> headersHorizantal =
                XlHelper.GetXlHeadersTableHorizontal(Globals.wShe);
            Dictionary<string, int> headersVertical =
                XlHelper.GetXlHeadersTableVertical(Globals.wShe);
            foreach (var item in headersVertical)
            {
                //bidimensional[1][1] = "a";
                string ColumnaVertical = item.Key;
                if (headersHorizantal.ContainsKey(ColumnaVertical))
                {
                    int columnaHorizontal = item.Value;
                    int ColumnaValueVertical = GetHorizontalValue(headersHorizantal, ColumnaVertical);
                    Range celdaVeritcal = Globals.wShe.Cells[1, columnaHorizontal];
                    celdaVeritcal.Select();
                    Range celdaHorizontal = Globals.wShe.Cells[ColumnaValueVertical, 1];
                    celdaHorizontal.Select();
                    FillExcel(ColumnaValueVertical, columnaHorizontal, ColumnaVertical);
                }

            }

        }

        public void CloseExcel()
        {
            try
            {
                string date = "";
                string time = "";
                string cadena = "";
                date = DateTime.Now.ToString("dd-mm-yy");
                time = DateTime.Now.ToString("mm-ss-hh");
                cadena = $"{date}{time}";
                Globals.xlWk.SaveAs(cadena);
                Globals.xlWk.Close();
                Globals.xlApp.Quit();
                Globals.xlApp = null;

            }
            catch (Exception ex)
            {

            }

        }
        public void createSheet(string name)
        {
            try
            {
                Globals.wShe = Globals.xlWk.Sheets.Add();
                Globals.wShe.Name = name;

            }
            catch (Exception ex)
            {

            }
        }
    }
}
