using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelService
{
    public interface IExcel
    {
        void OpenWorkBook(string path);
        void CloseWorkBook();
        string ReadValue(int y, int x);
        void SaveWorkBook(string path);
    }
}
