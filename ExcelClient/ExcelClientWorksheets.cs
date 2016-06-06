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
        private Excel.Worksheet GetWorkSheetWithName(Excel.Workbook workBook, string WorksheetName, bool CreateNewIfNotFound = true)
        {
            foreach (var wksht in workBook.Worksheets)
            {
                Excel.Worksheet thisWorksheet = (Excel.Worksheet)wksht;
                if (thisWorksheet != null)
                {
                    if (thisWorksheet.Name == WorksheetName)
                    {
                        return thisWorksheet;
                    }
                }
            }
            if (CreateNewIfNotFound == true)
            {

                Excel.Worksheet newWorksheet = (Excel.Worksheet)workBook.Worksheets.Add();
                newWorksheet.Name = WorksheetName;
                return newWorksheet;
            }
            throw new Exception("Worksheet (tab) with the specified name not found in Excel file for input values. Please check input.");
        }

        public List<string> GetWorksheetsWithPrefix(string WorkbookPath, string prefix)
        {

            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook xlwkbook = excelApplication.Workbooks.Open(WorkbookPath);
            List<string> wkshts = GetWorksheetsWithPrefix(xlwkbook, prefix);


            excelApplication.Quit();
            excelApplication.Dispose();
            return wkshts;
        }

        private List<string> GetWorksheetsWithPrefix(Excel.Workbook xlwkbook,  string prefix)
        {
            List<string> wkshts = new List<string>();

            foreach (var wksht in xlwkbook.Worksheets)
            {
                Excel.Worksheet thisWorksheet = (Excel.Worksheet)wksht;
                if (thisWorksheet != null)
                {
                    if (thisWorksheet.Name.StartsWith(prefix))
                    {
                        wkshts.Add(thisWorksheet.Name);
                    }
                }
            }

            return wkshts;
        }
    }
}
