using DataAccess;
using Entity;
using ExcelService;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Logic
{
    public class Core : IDisposable
    {
        private static readonly string comtecFile = "1.xls";

        private static readonly string currentLocation = 
            System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        private static readonly string comtecFilePath =
            Path.Combine(System.IO.Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location), comtecFile);

        Excel excel = new Excel();
        DataProvider provider = new DataProvider();

        List<Item> comtecList;
        List<Item> ammList;
        List<Item> mergedList;
        List<string> storesList;

        public void Prepare()
        {
            comtecList = new List<Item>();
            ammList = new List<Item>();
            mergedList = new List<Item>();
            storesList = new List<string>();
        }

        public void SetStores(List<string> stores)
        {
            storesList = stores;
        }

        public List<string> GetAmmStores()
        {
            return provider.GetStores();
        }


        public void SavaAndExit(string s)
        {
            string path = 
                Path.Combine(currentLocation, s);
            excel.SaveWorkBook(path);
            comtecList = null;
            ammList = null;
            mergedList = null;
            storesList = null;
            p = 1;
            excel.Dispose();
        }

        public void PrintResult(bool consoleOut, bool party)
        {
            excel.AddWorkbook();
            int gg = 1;
            excel.Write(gg, 1, !party ? "мх" : "партия");
            excel.Write(gg, 2, "нп");
            excel.Write(gg, 3, "комтех");
            excel.Write(gg, 4, "амм");

            foreach (var u in mergedList)
            {
                gg++;
                excel.Write(gg, 1, !party ? u.store : u.party);
                excel.Write(gg, 2, u.name);
                excel.Write(gg, 3, u.cQuantity.ToString("0.#####"));
                excel.Write(gg, 4, u.aQuantity.ToString("0.#####"));
                if (consoleOut)
                    ConsoleOut(gg);
            }
        }

        public int MergeData(bool party)
        {
            mergedList = ammList
                .Join(
                    comtecList,
                    a => new
                    {
                        Store = !party ? a.store : null,
                        Name = a.name,
                        Party = party ? a.party : null
                    },
                    c => new
                    {
                        Store = !party ? c.store : null,
                        Name = c.name,
                        Party = party ? c.party : null
                    },
                    (a, c) => new
                    {
                        union = a,
                        ii = c
                    })
                .Where(
                    w => w.union.aQuantity != w.ii.cQuantity)
                .Select(
                    s => new Item()
                    {
                        name = s.union.name,
                        store = !party ? s.union.store : null,
                        aQuantity = s.union.aQuantity,
                        cQuantity = s.ii.cQuantity,
                        party = party ? s.union.party : null
                    })
                    .ToList();

            return mergedList.Count;
        }

        public List<string> GetComtecStores()
        {
            storesList = comtecList.Select(s => s.store).Distinct().ToList();
            return storesList;
        }


        public int ConnectToAMMDB(bool party)
        {
            ammList = provider.GetQuery(storesList, party);
            return ammList.Count;
        }

        public void ReadComtecFile(bool consoleOut, bool party)
        {
            excel.OpenWorkBook(comtecFilePath);
            int lastRow = excel.LastRow();
            for (int i = 2; i <= lastRow; i++)
            {
                comtecList.Add(new Item
                {
                    store = !party ? excel.ReadValue(i, 1) : null,
                    party = party ? excel.ReadValue(i, 1) : null,
                    name = excel.ReadValue(i, 2),
                    cQuantity = Convert.ToDecimal(excel.ReadValue(i, 3))
                });
                if (consoleOut)
                    ConsoleOut(i);
            }
            excel.CloseWorkBook();
        }

        int p = 1;
        private void ConsoleOut(int i)
        {
            string s;
            s = "обработано строк: " + i;
            Console.Write("".PadRight(p));
            Console.Write("\r");
            Console.Write(s);
            Console.Write("\r");
            p = s.Length;
        }

        public void Dispose()
        {
            excel.Dispose();
        }
    }
}
