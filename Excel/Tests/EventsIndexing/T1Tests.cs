#if DEBUG

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
            var man = new ExcelManager(Path.Combine("Mw/Excel/Tests/Recon", "00_База_2018_10_01.xls"));
            man.Read();
            
            // group by year

            foreach (HSSFRow row in man.Records)
            {
                var index = man.GetValue(row, "Индекс");
            }

            // fill collumns
            // make symmetry
            // output
        }
    }
}
#endif
