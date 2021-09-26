using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DailyReportApp
{
    class ExcelUtils
    {
        public static void AppendRows(string path, string[,] rows)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook sheet = excel.Workbooks.Open(path);
            Excel.Worksheet x = excel.ActiveSheet as Excel.Worksheet;
            Excel.Range userRange = x.UsedRange;
            int countRecords = userRange.Rows.Count;
            int startIndex = countRecords + 1;
            Console.WriteLine(rows.GetLength(0));
            Console.WriteLine(rows.GetLength(1));
            for (int rowIndex = 0; rowIndex < rows.GetLength(0); rowIndex++)
            {
                for (int columnIndex = 0; columnIndex < rows.GetLength(1); columnIndex++)
                {
                    x.Cells[startIndex + rowIndex, columnIndex + 1] = rows[rowIndex, columnIndex];
                }
            }
                
            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
        }
    }
}
