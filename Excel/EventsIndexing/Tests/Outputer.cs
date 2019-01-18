using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Routines.Excel.EventsIndexing.Tests
{
    public abstract class Outputer
    {
        protected const string Folder = "Excel/EventsIndexing/Tests";

        public abstract string FolderOut { get; }

        public static string OutputConsole(Dictionary<string, Dictionary<string, int>> rows)
        {
            return new FullDbLauncher().Output(rows, null, OutputType.Console);
        }

        public string Output(
            Dictionary<string, Dictionary<string, int>> rows,
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
                    cell.SetCellValue(string.Empty);
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