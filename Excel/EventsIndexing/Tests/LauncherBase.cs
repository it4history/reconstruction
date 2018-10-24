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
        protected const string Folder = "Excel/EventsIndexing/Tests";

        private const string FilterString20 = "яш кя ля пс ую щб рю бж фг фю сг цх ак цз ск фм аю ип нв фп";
        private const string FilterString40 = FilterString20 + " ай ба не ду вы сю бв щд шы мч хе щг ти цт лх сх гя лс пь вз";
        private static readonly string[] Filter = FilterString40.Split(' ');

        /// <summary>
        /// for column Индекс
        /// </summary>
        public const char IndicesSeparator = ',';

        /// <summary>
        /// XLS format is faster for reading
        /// </summary>
        public abstract string FileNameIn { get; }

        public abstract string FolderOut { get; }

        /// <summary>
        /// big files are slow in Xlsx for writing
        /// </summary>
        public virtual OutputType FileOutputType { get { return OutputType.Xlsx; } }

        public virtual bool ReadLegend { get { return false; } }

#region http://hist.tk/hw/EventsIndexing
        /// <summary>
        /// как часто разные события встречаются в одном и том же году
        /// </summary>
        [Test]
        public void Do()
        {
            var eventsMan = new ExcelManager(Path.Combine(Folder, FileNameIn));
            DoWithShift(GroupEventsByYear(eventsMan), 0);
            DoWithShift(GroupEventsByYear(eventsMan, null, true), 0);

            if (ReadLegend)
            {
                var man = new ExcelManager(eventsMan, 1); // Легенда sheet
                var legends = GetLegend(man, "Подгруппа");
                DoWithShift(GroupEventsByYear(eventsMan, legends), 0);
                DoWithShift(GroupEventsByYear(eventsMan, legends, true), 0);
            }
        }

        protected static Dictionary<string, string> GetLegend(ExcelManager man, string groupColumn)
        {
            var legends = new Dictionary<string, string>();
            foreach (HSSFRow row in man.Records)
            {
                var index = row.Cells[0].StringCellValue;
                var legend = man.GetValue(row, groupColumn);
                if (!string.IsNullOrEmpty(legend))
                {
                    legends.Add(index, legend.ToLower());
                }
            }
            return legends;
        }


        /// <summary>
        /// совпадения со сдвигом в 1-5 год
        /// </summary>
        [Test]
        public void DoWithShifts()
        {
            var man = new ExcelManager(Path.Combine(Folder, FileNameIn));
            var byYears = GroupEventsByYear(man);
            DoWithShift(byYears, 1);
            DoWithShift(byYears, 2);
            DoWithShift(byYears, 3);
            DoWithShift(byYears, 4);
            DoWithShift(byYears, 5);
        }

        private void DoWithShift(Graphes byYears, int shift)
        {
            var s = shift == 0 ? null : string.Format("_сдвиг_{0}год", shift);
            var rows = EventtypesByYears.Do(byYears, shift);
            Output(
                rows,
                string.Format("сводная{0}{1}{2}",
                    byYears.TwoAndMoreEventtypes ? "_filter2_" : null,
                    byYears.Legends != null ? "Группирована" : null,
                    s),
                FileOutputType);
            // Output(rows, "проценты" + s, true, null, true);
            if (!byYears.TwoAndMoreEventtypes && byYears.Legends == null)
                Output(
                    EventtypesByYears.Do(byYears, shift), // 822, 1852),
                    "сводная" + Filter.Length + s, OutputType.Csv, Filter);
        }
        #endregion

        protected virtual Graphes GroupEventsByYear(
            ExcelManager man, 
            Dictionary<string, string> legends = null,
            bool twoAndMoreEventtypes = false)
        {
            var result = new Graphes(legends, twoAndMoreEventtypes);
            foreach (HSSFRow row in man.Records)
            {
                var indices = man.GetValue(row, "Индекс");
                if (!string.IsNullOrEmpty(indices)
                    && (!twoAndMoreEventtypes
                        || indices.Split(IndicesSeparator).Count(i => !string.IsNullOrEmpty(i)) >= 2))
                {
                    if (legends != null)
                    {
                        var indicesFromLegend = new List<string>();
                        foreach (var index in indices.Split(IndicesSeparator))
                        {
                            indicesFromLegend.Add(legends.ContainsKey(index) ? legends[index] : index);
                        }
                        indices = indicesFromLegend.Aggregate((a, b) => a + IndicesSeparator + b);
                    }
                    var eventOrYear = twoAndMoreEventtypes ? row.RowNum : int.Parse(man.GetValue(row, "-99000"));
                    result.Add(eventOrYear, indices);
                }
            }
            return result;
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

            Console.WriteLine("graph {1} has {0} nodes", sorted.Count(), name);

            var path = GetOutputPath(name, type);
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

        protected string GetOutputPath(string name, OutputType type)
        {
            var folder = string.Format("../../{0}/{1}", Folder, FolderOut);
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
            return Path.Combine(folder, name + "." + type.ToString().ToLower());
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
