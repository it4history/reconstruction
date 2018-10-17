#if DEBUG

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Logy.Api.Mw.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NUnit.Framework;

namespace Routines.Excel.EventsIndexing
{
    public class T1Tests
    {
        const string Folder = "Excel/EventsIndexing/Tests";

        private static readonly string[] Filter
            = "яш кя ля пс ую щб рю бж фг фю сг цх ак цз ск фм аю ип нв".Split(' ');
            // " фп ай ба не ду вы сю бв щд шы мч хе щг ти цт лх сх гя лс пь"

        /// <summary>
        /// как часто разные события встречаются в одном и том же году
        /// </summary>
        [Test]
        public void Do()
        {
            var byYears = Read();
            DoWithShift(byYears, 0);
        }


        /// <summary>
        /// совпадения со сдвигом в 1-5 год
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
                EventtypesByYears.Do(byYears, null, shift, 822, 1852),
                "сводная20" + s, OutputType.Csv, Filter); // suffix 20 because Filter had 20 indices
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

        public static string Output(Dictionary<string, Dictionary<string, int>> rows,
                           string name,
                           OutputType type = OutputType.Xlsx,
                           string[] outOrder = null,
                           bool percents = false)
        {
            var file = string.Format("../../{0}/out/{1}.{2}", Folder, name, type.ToString().ToLower());
            var sorted = outOrder ?? (IEnumerable<string>)rows.Keys
                .Where(s => rows[s].Count > 0 || !ColumnIsEmpty(rows, s))
                .OrderBy(s => s);
            if (type == OutputType.Xlsx)
            {
                using (var stream = new FileStream(file, FileMode.Create, FileAccess.Write))
                {
                    var wb = new XSSFWorkbook();
                    var sheet = wb.CreateSheet(name);
                    var cH = wb.GetCreationHelper();

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

                    var diagonalStyle = wb.CreateCellStyle();
                    diagonalStyle.FillBackgroundColor = IndexedColors.Yellow.Index;
                    diagonalStyle.FillPattern = FillPattern.LeastDots;
                    foreach (var row in sorted)
                    {
                        erow = sheet.CreateRow(erowN++);
                        ecellN = 0;
                        cell = erow.CreateCell(ecellN++);
                        cell.SetCellValue(row);

                        var cols = rows[row];
                        foreach (var col in sorted)
                        {
                            cell = erow.CreateCell(ecellN++);
                            cell.SetCellValue(cols.ContainsKey(col) ? cols[col] : 0);
                            if (row == col)
                                cell.CellStyle = diagonalStyle;
                        }
                    }
                    wb.Write(stream);
                }
            }
            else
            {
                var sb = new StringBuilder();
                sb.Append("$, ");
                foreach (var col in sorted)
                {
                    sb.Append(col + ", ");
                }
                sb.AppendLine();
                foreach (var row in sorted)
                {
                    sb.Append(row + ", ");
                    var cols = rows[row];
                    foreach (var col in sorted)
                    {
                        sb.Append(cols.ContainsKey(col) ? cols[col] : 0);
                        sb.Append(", ");
                    }
                    sb.AppendLine();
                }
                if (type == OutputType.Csv)
                {
                    File.WriteAllText(file, sb.ToString());
                }
                return sb.ToString();
            }
            return null;
        }

        private static bool ColumnIsEmpty(Dictionary<string, Dictionary<string, int>> rows, string column)
        {
            foreach (var pair in rows.Values)
            {
                // pair does not contain 0 counts
                if (pair.ContainsKey(column))
                {
                    return false;
                }
            }
            return true;
        }
    }

    public enum OutputType
    {
        Xlsx,
        Csv,
        Console
    }
}
#endif
