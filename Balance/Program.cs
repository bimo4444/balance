using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Logic;

namespace Balance
{
    class Program
    {
        static Core core = new Core();
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            AppDomain.CurrentDomain.ProcessExit += OnExit;
            DateTime dt = DateTime.Now;

            Console.WriteLine("****************************************");
            Console.WriteLine("* Сравнение остатков АММ - Комтех v1.0 *");
            Console.WriteLine("****************************************");
            Console.WriteLine();
            Console.WriteLine("Расположите в папке с программой выгрузку из Комтеха на " + dt.Date.AddDays(-1).ToShortDateString() + 
                " включительно, с именем 1.xls и введите латинскую 's' для обычного сравнения или латинскую 'p' для сравнения с учётом партий. Нажмите Enter для выхода из программы");
            Loop();
            
        }
        private static void Loop()
        {
            string s;
            while (true)
            {
                try
                {
                    Console.WriteLine();
                    s = Console.ReadLine();
                    switch (s)
                    {
                        case "s":
                            Action();
                            break;
                        case "p":
                            Action(true);
                            break;
                        case "":
                            Environment.Exit(0);
                            break;
                    }
                }
                catch (Exception x)
                {
                    Console.WriteLine("Произошла ошибка:");
                    Console.WriteLine(x.Message);
                }
            }
        }

        private static List<string> GetStores()
        {
            List<string> result;
            Console.WriteLine("получение списка мест хранения...");
            var ammStores = core.GetAmmStores();
            int i = 0;
            var stores = ammStores.Select(s => new { id = i++.ToString(), store = s }).ToList();
            foreach (var r in stores)
                Console.WriteLine(r.id + ". " + r.store);
            Console.WriteLine();
            Console.WriteLine("для продолжения введите номер интересующего маста хранения");
            while(true)
            {
                Console.WriteLine();
                string input = Console.ReadLine();
                result = stores.Where(w => w.id == input).Select(s => s.store).ToList();
                if (result != null && result.Count < 1)
                    Console.WriteLine("введено не верное значение");
                else
                    break;
            }
            Console.WriteLine();
            return result;
        }

        private static void Action(bool party = false)
        {
            List<string> stores;
            Console.WriteLine();
            core.Prepare();
            Console.WriteLine("чтение выгрузки из комтеха...");
            core.ReadComtecFile(consoleOut: true, party: party);
            Console.WriteLine();
            Console.WriteLine();

            if (party)
            {
                stores = GetStores();
                core.SetStores(stores);
            }
            else
            {
                Console.WriteLine("получение списка мест хранения...");
                stores = core.GetComtecStores();
                foreach (var v in stores)
                    Console.WriteLine(v);
                Console.WriteLine();
            }
            
            Console.WriteLine("подключение к базе...");
            int ammCount = core.ConnectToAMMDB(party);
            Console.WriteLine("получено строк из амм: " + ammCount);
            Console.WriteLine();

            Console.WriteLine("сравнение информации...");
            int mergedCount = core.MergeData(party);
            Console.WriteLine("не совпадений: " + mergedCount);
            Console.WriteLine();

            Console.WriteLine("запись в книгу...");
            core.PrintResult(consoleOut: true, party: party);
            Console.WriteLine();

            string dt = DateTime.Now.ToString("dd.MM.yyyy HH-mm-ss") + ".xlsx";
            Console.WriteLine("сохранение...");
            core.SavaAndExit(dt);
            Console.WriteLine();

            Console.WriteLine("Файл отчёта сохранён в папке с программой под именем " + dt);
            Console.WriteLine();

            Console.WriteLine("Для выхода из программы нажмите Enter, для повторного сравнения введите 's' или 'p' для сравнения с учётом партий");
            Loop();
        }

        private static void OnExit(object sender, EventArgs e)
        {
            core.Dispose();
        }
    }
}
