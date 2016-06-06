using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using NetOffice;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace Udilovich.ExcelClient
{
    public partial class ExcelClientNetOffice
    {

        public void DumpMultipleValueArrays(List<ExcelOutputEntry> DataOutputEntries, string ExcelOutputPath)
        {

            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook xlwkbook = excelApplication.Workbooks.Open(ExcelOutputPath);

            foreach (var entry in DataOutputEntries)
            {
                DumpArrayOfValues(entry.Values, entry.StartRow, entry.StartColumn, xlwkbook, entry.Worksheet);
            }

            xlwkbook.Save();
            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();
        }


        private void DumpArrayOfValues(string[,] Values, int StartRow, int StartColumn, Excel.Workbook xlwkbook, string WorksheetName)
        {

            Excel.Worksheet xlsheet = GetWorkSheetWithName(xlwkbook, WorksheetName);

            for (int i = 0; i < Values.GetLength(0); i++)
            {
                for (int j = 0; j < Values.GetLength(1); j++)
                {
                    int cRow = StartRow + i;
                    int cColumn = StartColumn + j;
                    Excel.Range CurrentRange = (Excel.Range)xlsheet.Cells[cRow, cColumn];
                    CurrentRange.Cells.Value = Values[i, j];
                }
            }


        }

    }
}
