using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using X = Microsoft.Office.Interop.Excel;

namespace ExcelService
{
    public class Excel : IExcel, IDisposable
    {
        X.Application excelApp;
        X.Workbook workBook;
        X.Worksheet workSheet;

        public void OpenWorkBook(string path)
        {
            excelApp = new X.Application();
            excelApp.Visible = false;
            excelApp.Interactive = false;
            excelApp.DisplayAlerts = false;
            workBook = excelApp.Workbooks.Open(path);
            Convert(path);
        }

        private void Convert(string path)
        {
            SaveWorkBook(path + "x");
            workBook = excelApp.Workbooks.Open(path + "x");
            workSheet = workBook.ActiveSheet;
        }

        public void AddWorkbook()
        {
            excelApp.Workbooks.Add();
            workBook = excelApp.ActiveWorkbook;
            workSheet = workBook.ActiveSheet;
        }
        public void Write(int y, int x, string s)
        {
            workSheet.Cells[y, x] = s;
        }
        public void CloseWorkBook()
        {
            workBook.Close();
        }

        public int LastRow()
        {
            return workSheet.Cells.SpecialCells(X.XlCellType.xlCellTypeLastCell).Row;
        }

        public string ReadValue(int y, int x)
        {
            return ((X.Range)workSheet.Cells[y, x]).Text ?? "";
        }

        public void SaveWorkBook(string path)
        {
            workBook.SaveAs(path, X.XlFileFormat.xlOpenXMLWorkbook, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing,
                X.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            CloseWorkBook();
        }

        public void Dispose()
        {
            if (excelApp == null)
                return;
            excelApp.Visible = true;
            excelApp.Interactive = true;
            excelApp.DisplayAlerts = true;
            excelApp.Quit();
            excelApp = null;
        }
    }
}
