using Entity;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess
{
    public class DataProvider
    {
        public List<string> GetStores()
        {
            using (DataClasses1DataContext l = new DataClasses1DataContext())
            {
                l.CommandTimeout = 300;
                return l.viewForComtecBalanceChecks
                    .Select(s => s.storeName)
                    .Distinct()
                    .OrderBy(o => o)
                    .ToList();
            }
        }

        public List<Item> GetQuery(List<string> storeNames, bool party)
        {
            using (DataClasses1DataContext l = new DataClasses1DataContext())
            {
                l.CommandTimeout = 300;
                List<Item> list = l.viewForComtecBalanceChecks
                    .Where(
                        w => storeNames.Contains(w.storeName)
                        && w.dt < DateTime.Now.Date)
                    .GroupBy(
                    g => new { StoreName = g.storeName, Np = g.np, Party = party ? g.party : null })
                    .Select(
                        s => new Item
                        {
                            store = s.Key.StoreName,
                            name = s.Key.Np,
                            aQuantity = s.Sum(ss => ss.qty) ?? (decimal)0.0,
                            party = party ? s.Key.Party : null
                        })
                    .ToList();

                return !party ? list : list.Where(w => w.party != null).ToList();
            }
        }
    }

}
