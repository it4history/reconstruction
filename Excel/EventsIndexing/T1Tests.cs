#if DEBUG

using System.Collections.Generic;
using System.IO;
using System.Linq;
using Logy.Api.Mw.Excel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NUnit.Framework;

namespace Routines.Excel.EventsIndexing
{
    public class T1Tests
    {
        const string Folder = "Excel/EventsIndexing/Tests";
        const string Filter = "яш кя ля пс ую щб рю бж фг фю сг цх ак цз ск фм аю ип нв";

        /// <summary>
        /// как часто разные события встречаются в одном и том же году году
        /// </summary>
        [Test]
        public void Do()
        {
            var byYears = Read();
            DoWithShift(byYears, 0);
        }

        /// <summary>
        /// совпадения со сдвигом в год, два, три
        /// </summary>
        [Test]
        public void DoWithShifts()
        {
            var byYears = Read();
            DoWithShift(byYears, 1);
            DoWithShift(byYears, 2);
            DoWithShift(byYears, 3);
            DoWithShift(byYears, 4);
            DoWithShift(byYears, 5);
        }

        private void DoWithShift(Dictionary<int, string> byYears, int shift)
        {
            var s = shift == 0 ? null : string.Format("_сдвиг_{0}год", shift);
            var rows = EventtypesByYears.Do(byYears, null, shift);
            Output(rows, "сводная" + s);
            // Output(rows, "проценты" + s, true, null, true);
            Output(
                EventtypesByYears.Do(byYears, Filter, shift, 822, 1852),
                "сводная20" + s, false, Filter);
        }

        public Dictionary<int, string> Read()
        {
            var man = new ExcelManager(Path.Combine(Folder, "00_База_2018_10_01.xls"));
            man.Read();

            // group events by year
            var byYears = new Dictionary<int, string>();
            foreach (HSSFRow row in man.Records)
            {
                var indices = man.GetValue(row, "Индекс");
                if (!string.IsNullOrEmpty(indices))
                {
                    var year = int.Parse(man.GetValue(row, "-99000"));
                    byYears[year] = (byYears.ContainsKey(year) ? byYears[year] : null) + indices;
                }
            }
            return byYears;
        }

        public void Output(Dictionary<string, Dictionary<string, int>> rows,
                           string name,
                           bool inXlsx = true,
                           string outOrder = null,
                           bool percents = false)
        {
            var file = string.Format("../../{0}/out/{1}.{2}", Folder, name, inXlsx ? "xlsx" : "csv");
            var sorted = outOrder != null 
                ? outOrder.Split(' ') 
                : (IEnumerable<string>)rows.Keys.Where(s => rows[s].Count > 0).OrderBy(s => s);
            if (inXlsx)
            {
                using (var stream = new FileStream(file, FileMode.Create, FileAccess.Write))
                {
                    var wb = new XSSFWorkbook();
                    var sheet = wb.CreateSheet(name);
                    // var cH = wb.GetCreationHelper();

                    var erowN = 0;
                    var ecellN = 0;
                    var erow = sheet.CreateRow(erowN++);
                    var cell = erow.CreateCell(ecellN++);
                    cell.SetCellValue("");
                    foreach (var col in sorted)
                    {
                        cell = erow.CreateCell(ecellN++);
                        cell.SetCellValue(col);
                    }

                    foreach (var row in sorted)
                    {
                        erow = sheet.CreateRow(erowN++);
                        ecellN = 0;
                        cell = erow.CreateCell(ecellN++);
                        cell.SetCellValue(row);

                        var cols = rows[row];
                        foreach (var col in sorted)
                        {
                            if (cols.ContainsKey(col))
                            {
                                cell = erow.CreateCell(ecellN++);
                                cell.SetCellValue(cols[col]);
                            }
                        }
                    }
                    wb.Write(stream);
                }
            }
            else
            {
                var console = File.CreateText(file);
                console.Write(",");
                foreach (var col in sorted)
                {
                    console.Write(col + ", ");
                }
                console.WriteLine();
                foreach (var row in sorted)
                {
                    console.Write(row + ", ");
                    var cols = rows[row];
                    foreach (var col in sorted)
                    {
                        if (cols.ContainsKey(col))
                        {
                            console.Write(cols[col]);
                        }
                        console.Write(", ");
                    }
                    console.WriteLine();
                }
                console.Close();
            }
        }
    }
}
#endif
