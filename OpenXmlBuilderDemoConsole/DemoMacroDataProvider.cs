using OpenXmlBuilder;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace OpenXmlBuilderDemoConsole
{
    internal class DemoMacroDataProvider : IMacroDataProvider
    {
        private readonly Dictionary<string, IEnumerable<object>> dict = new Dictionary<string, IEnumerable<object>>();

        public IEnumerable<object> GetData(string aKey)
        {
            if (!dict.ContainsKey(aKey))
            {
                return new List<object>();
            }

            return dict[aKey];
        }

        public void init()
        {
            var list1 = Enumerable.Range(0, 100).Select(x => new List1Item
            {
                Product     = $"Name {x}",
                Price       = new Random().Next(100) + (decimal?)new Random().NextDouble(),
                Qnt         = new Random().Next(10),
                SaleDate    = new DateTime(2021, 1, 1).AddDays(new Random().Next(336))
            });
            dict.Add("list1", list1);

            var list2 = new List<List1Item> { new List1Item { Qnt = list1.Sum(x => x.Qnt), Price = list1.Sum(x => x.Price) } };
            dict.Add("list1Total", list2);

            var list0 = new List<List0Item> { new List0Item { DocumentNo = "777", ParentDocumentDate = DateTime.Now, ParentDocumentNo = "333" } };
            dict.Add("list0", list0);
        }
        
        public class List0Item
        {
            public string    DocumentNo         { get; set; }
            public string    ParentDocumentNo   { get; set; }
            public DateTime? ParentDocumentDate { get; set; }
        }

        public class List1Item
        { 
            public string   Product     { get; set; }
            public string   PhotoPath
            {
                get
                {
                    return $"{Directory.GetCurrentDirectory()}\\Source\\ms-.net-framework.jpg";
                }
                set { }
            }
            public decimal? Price       { get; set; }
            public int?     Qnt         { get; set; }
            public DateTime? SaleDate   { get; set; }
            public decimal? Summ => Qnt * Price;
        }
    }
}
