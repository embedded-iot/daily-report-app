using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace DailyReportApp
{
    static class ExcelUtils
    {
        public static void WriteArray<T>(this _Worksheet sheet, int startRow, int startColumn, T[,] array)
        {
            var row = array.GetLength(0);
            var col = array.GetLength(1);
            Range c1 = (Range)sheet.Cells[startRow, startColumn];
            Range c2 = (Range)sheet.Cells[startRow + row - 1, startColumn + col - 1];
            Range range = sheet.Range[c1, c2];
            string[,] data = new string[row, col];
            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < col; j++)
                {
                    data[i, j] = "232";
                }
            }
            range.NumberFormat = "0.0";
            range.Value = data;
            //range.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, array);
        }

        public static void AppendRowsNew(string excelPath, int sheetIndex, string[,] rows, bool isHasBorder)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook sheet = excel.Workbooks.Open(excelPath);
            //Excel.Worksheet x = excel.ActiveSheet as Excel.Worksheet;
            Excel.Worksheet x = (Excel.Worksheet)sheet.Worksheets[sheetIndex];
            Excel.Range userRange = x.UsedRange;
            int countRecords = userRange.Rows.Count;
            int startIndex = countRecords + 1;
            x.WriteArray(startIndex, 1, rows);
            Excel.Range newUserRange = x.UsedRange;
            if (isHasBorder)
            {
                newUserRange.Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            }
            sheet.Save();
            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
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
            for (int rowIndex = 0; rowIndex < rows.GetLength(0); rowIndex++)
            {
                //Console.WriteLine("Line " + rowIndex + " - at " + DateTime.Now);
                for (int columnIndex = 0; columnIndex < rows.GetLength(1); columnIndex++)
                {
                    x.Cells[startIndex + rowIndex, columnIndex + 1] = rows[rowIndex, columnIndex]; // lỗi đấy
                    
                    if (isHasBorder)
                    {
                        x.Cells[startIndex + rowIndex, columnIndex + 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;   
                    }
                    
                }
            }
                
            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
        }

        //public static string[] ReadRow(string path, int sheetIndex, int rowIndex)
        //{
        //    Excel.Application excel = new Excel.Application();
        //    Excel.Workbook sheet = excel.Workbooks.Open(path);
        //    Excel.Worksheet x = (Excel.Worksheet)sheet.Worksheets[sheetIndex];
        //    Excel.Range userRange = x.UsedRange;
            
        //    int countColumns = userRange.Columns.Count;
        //    string[] data = new string[countColumns];
        //    int columnIndex = 1;
        //    while (columnIndex <= countColumns)
        //    {
        //        if (x.Cells[rowIndex, columnIndex] != null && x.Cells[rowIndex, columnIndex].Value2 != null)
        //        {
        //            data[columnIndex-1] = x.Cells[rowIndex, columnIndex].Value2.ToString();
        //        }
        //        columnIndex++;
        //    }

        //    sheet.Close(true, Type.Missing, Type.Missing);
        //    excel.Quit();
        //    return data;
        //}
    }
}
