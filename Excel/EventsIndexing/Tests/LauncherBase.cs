#if DEBUG

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Logy.Api.Mw.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NUnit.Framework;

namespace Routines.Excel.EventsIndexing.Tests
{
    public abstract class LauncherBase
    {
        const string Folder = "Excel/EventsIndexing/Tests";

        private const string FilterString20 = "яш кя ля пс ую щб рю бж фг фю сг цх ак цз ск фм аю ип нв";
        private const string FilterString40 = FilterString20 + " фп ай ба не ду вы сю бв щд шы мч хе щг ти цт лх сх гя лс пь";
        private static readonly string[] Filter = FilterString40.Split(' ');

        /// <summary>
        /// XLS format is faster for reading
        /// </summary>
        public abstract string FileNameIn { get; }

        public abstract string FolderOut { get; }

        /// <summary>
        /// big files are slow in Xlsx for writing
        /// </summary>
        public virtual OutputType FileOutputType { get { return OutputType.Xlsx; } }

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
            Output(rows, "сводная" + s, FileOutputType);
            // Output(rows, "проценты" + s, true, null, true);
            Output(
                EventtypesByYears.Do(byYears, null, shift), // 822, 1852),
                "сводная40" + s, OutputType.Csv, Filter); // suffix shows how many indices
        }

        public Dictionary<int, string> Read()
        {
            var man = new ExcelManager(Path.Combine(Folder, FileNameIn));
            man.Read();
            return GroupEventsByYear(man);
        }

        protected virtual Dictionary<int, string> GroupEventsByYear(ExcelManager man)
        {
            var byYears = new Dictionary<int, string>();
            foreach (HSSFRow row in man.Records)
            {
                var indices = man.GetValue(row, "Индекс");
                if (!string.IsNullOrEmpty(indices))
                {
                    var year = int.Parse(man.GetValue(row, "-99000"));
                    byYears[year] = (byYears.ContainsKey(year) ? byYears[year] : null) + indices + ",";
                }
            }
            return byYears;
        }

        public static string OutputConsole(Dictionary<string, Dictionary<string, int>> rows)
        {
            return new FullDbLauncher().Output(rows, null, OutputType.Console);
        }

        public string Output(Dictionary<string, Dictionary<string, int>> rows,
                           string name,
                           OutputType type = OutputType.Xlsx,
                           string[] outOrder = null,
                           bool percents = false)
        {
            var sorted = outOrder ?? (IEnumerable<string>)rows.Keys
                .Where(s => rows[s].Count > 0 || !ColumnIsEmpty(rows, s))
                .OrderBy(s => s).ToArray();

            Console.WriteLine("graph has {0} nodes", sorted.Count());

            var folder = string.Format("../../{0}/{1}", Folder, FolderOut);
            var path = Path.Combine(folder, name + "." + type.ToString().ToLower());
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
            if (type == OutputType.Xlsx)
            {
                using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write))
                {
                    var wb = new XSSFWorkbook();
                    var sheet = wb.CreateSheet(name);

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
                    stream.Close();
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
                    File.WriteAllText(path, sb.ToString());
                }
                else
                {
                    // OutOfMemoryException for graph with 7000 nodes
                    return sb.ToString();
                }
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
}
#endif
