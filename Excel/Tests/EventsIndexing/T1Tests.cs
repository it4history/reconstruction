#if DEBUG

using System.Collections.Generic;
using System.IO;
using NPOI.HSSF.UserModel;
using NUnit.Framework;

namespace Logy.Api.Mw.Excel.Tests.EventsIndexing
{
    public class T1Tests
    {
        [Test]
        public void Read()
        {
            var man = new ExcelManager(Path.Combine("Excel/Tests/EventsIndexing", "00_База_2018_10_01.xls"));
            man.Read();
            
            // group by year
            var byYears = new Dictionary<short, string>();
            foreach (HSSFRow row in man.Records)
            {
                var indices = man.GetValue(row, "Индекс");
                if (!string.IsNullOrEmpty(indices)){
                var year = man.GetYear(row);
                	foreach(var index in indices.Split(','))
                	{
                		
                	}
                }
            }

            // fill collumns
            // make symmetry
            // output
        }
    }
}
#endif
