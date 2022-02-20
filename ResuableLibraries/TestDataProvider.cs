using System;
using System.Collections;
using System.Runtime.InteropServices;
using xl = Microsoft.Office.Interop.Excel;

namespace SeleniumTestAutomationFramework.ResuableLibraries
{
    public class TestDataProvider
    {
        public xl.Application XlApp = null;
        public xl.Workbooks Workbooks = null;
        public xl.Workbook Workbook = null;
        public Hashtable Sheets;
        public string XlFilePath;

        public TestDataProvider(string XlFilePath)
        {
            this.XlFilePath = XlFilePath;
        }

        public void OpenExcel()
        {
            XlApp = new xl.Application();
            Workbooks = XlApp.Workbooks;
            Workbook = Workbooks.Open(XlFilePath);
            Sheets = new Hashtable();
            int count = 1;

            foreach (xl.Worksheet sheet in Workbook.Sheets)
            {
                Sheets[count] = sheet.Name;
                count++;
            }
        }

        public void CloseExcel()
        {
            Workbook.Close(false, XlFilePath, null);
            _ = Marshal.FinalReleaseComObject(Workbook);
            Workbook = null;

            Workbooks.Close();
            _ = Marshal.FinalReleaseComObject(Workbooks);
            Workbooks = null;

            XlApp.Quit();
            _ = Marshal.FinalReleaseComObject(XlApp);
            XlApp = null;
        }

        public string GetCellData(string sheetName, string colName, int rowNumber)
        {
            OpenExcel();

            string value = string.Empty;
            int sheetValue = 0;
            int colNumber = 0;

            if (Sheets.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in Sheets)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet Worksheet = Workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range Range = Worksheet.UsedRange;

                for (int i = 1; i <= Range.Columns.Count; i++)
                {
                    string colNameValue = Convert.ToString((Range.Cells[1, i] as xl.Range).Value2);

                    if (colNameValue.ToLower() == colName.ToLower())
                    {
                        colNumber = i;
                        break;
                    }
                }

                value = Convert.ToString((Range.Cells[rowNumber, colNumber] as xl.Range).Value2);
                _ = Marshal.FinalReleaseComObject(Worksheet);
            }
            CloseExcel();
            return value;
        }
    }
}

