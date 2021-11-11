using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace DailyReportApp
{
    static class ExcelUtility
    {
        public static void WriteArray<T>(this _Worksheet sheet, int startRow, int startColumn, T[,] array)
        {
            var row = array.GetLength(0);
            var col = array.GetLength(1);
            Range c1 = (Range)sheet.Cells[startRow, startColumn];
            Range c2 = (Range)sheet.Cells[startRow + row - 1, startColumn + col - 1];
            Range range = sheet.Range[c1, c2];
            range.Value = array;
        }

        public static bool ExportToExcel<T>(T[,] data, string path)
        {
            try
            {
                //Start Excel and get Application object.
                var oXl = new Application { Visible = false };

                //Get a new workbook.
                var oWb = (_Workbook)(oXl.Workbooks.Add(""));
                var oSheet = (_Worksheet)oWb.ActiveSheet;
                //oSheet.WriteArray(1, 1, bufferData1);

                oSheet.WriteArray(1, 1, data);

                oXl.Visible = false;
                oXl.UserControl = false;
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                oWb.SaveAs(path, XlFileFormat.xlWorkbookDefault, Type.Missing,
                    Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                oWb.Close(false);
                oXl.Quit();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }

            return true;
        }

        public static void AppendRows(string excelPath, int sheetIndex, string[,] rows, bool isHasBorder)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook sheet = excel.Workbooks.Open(excelPath);
            //Excel.Worksheet x = excel.ActiveSheet as Excel.Worksheet;
            Excel.Worksheet x = (Excel.Worksheet)sheet.Worksheets[sheetIndex];
            Excel.Range userRange = x.UsedRange;
            int countRecords = userRange.Rows.Count;
            int startIndex = countRecords + 1;
            Console.WriteLine("Start Write" + DateTime.Now);
            x.WriteArray(startIndex, 1, rows);
            sheet.Save();
            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
        }
    }
}
