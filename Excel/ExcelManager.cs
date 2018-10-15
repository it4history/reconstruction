using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Logy.Entities.Localization;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Logy.Api.Mw.Excel
{
    public class ExcelManager
    {
        private const string DescDatePattern = @"^(\d{2,4}(\s*[-—]\s*\d+)?)";
        private readonly string _path;
        private readonly bool _isXlsx;

        public ExcelManager(string path)
        {
            _path = path;
            _isXlsx = path.EndsWith("xlsx");
        }

        public IList Records { get; set; }
        public int RecordsCount { get { return Sheet.PhysicalNumberOfRows - 1; } }
        public List<string> Columns { get; set; }
        internal ISheet Sheet { get; set; }

        public string GetValue(IRow row, string column)
        {
            var cell = GetCell(row, column);
            return (cell != null) ? cell.ToString() : null;
        }

        public ICell GetCell(IRow row, string column)
        {
            var index = Columns.LastIndexOf(column);
            if (index == -1)
                index = Columns.LastIndexOf(column.ToLower());
            foreach (var cell in row.Cells)
            {
                if (cell.ColumnIndex == index)
                    return cell;
            }
            return null;
        }

        public void Read()
        {
            IWorkbook p;
            using (var fi = File.OpenRead(_path))
            {
                p = _isXlsx ? (IWorkbook)new XSSFWorkbook(fi) : new HSSFWorkbook(fi);
            }
            Sheet = p.GetSheetAt(0);
            var rows = Sheet.GetRowEnumerator();

            Records = _isXlsx ? (IList)new List<XSSFRow>() : new List<HSSFRow>();
            Columns = new List<string>();
            var firstRow = true;
            while (rows.MoveNext())
            {
                var row = (IRow)rows.Current;
                if (firstRow)
                {
                    foreach (var cell in row.Cells)
                    {
                        Columns.Add(cell.ToString());
                    }
                    firstRow = false;
                }
                else
                {
                    Records.Add(row);
                }
            }
        }

		internal string GetYear(IRow row)
		{
			return GetValue(
				row,
				JsonManager.GetJsonTranslation(ExcelFileColumns.Year));
		}
                
        internal static string[] GetYears(string title)
        {
            var s = Regex.Match(title, DescDatePattern + ".*").Groups[1].Value.Replace(" ", null);
            return string.IsNullOrEmpty(s) ? null : s.Split('-', '—');
        }

        internal static string TrimTitle(string title)
        {
            return Regex.Replace(TrimDescription(title), @"(\[.*\])", string.Empty);
        }

        internal static string TrimDescription(string title)
        {
            return title == null ? string.Empty : title.Trim(',', '.', ' ');
        }

        internal static string GetDescription(string title)
        {
            return TrimDescription(
                Regex.Match(title, DescDatePattern + @"?\s*(г+\.?,?о?[\w]{0,3})?\s(.+)").Groups[4].Value);
        }

        internal static string GetUrl(string source)
        {
            if (source != null)
            {
                var language = GetLanguage(source);
                source = source.Trim();
                if (language != null)
                {
                    return string.Format(
                        "[[wiki{0}:{1}]]",
                        language,
                        TrimDescription(source.Substring("Вікіпедія".Length))); // Википедия has same length
                }
                if (source.StartsWith("http"))
                    return source;
            }
            return null;
        }

        internal static string GetRowNumFromText(string text, string fileNameAsPropertyName)
        {
            return Regex.Match(text, string.Format(@"\[\[{0}-row::(\d+)\]\]", fileNameAsPropertyName)).Groups[1].Value;
        }

        internal static string GetLanguage(string source)
        {
            if (source != null)
            {
                if (source.StartsWith("Вікіпедія"))
                {
                    return "uk";
                }
                if (source.StartsWith("Википедия"))
                {
                    return "ru";
                }
            }
            return null;
        }
    }
}