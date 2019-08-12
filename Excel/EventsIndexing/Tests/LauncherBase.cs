#if DEBUG

using System.Collections.Generic;
using System.IO;
using System.Linq;
using Logy.Api.Mw.Excel;
using NPOI.HSSF.UserModel;
using NUnit.Framework;

namespace Routines.Excel.EventsIndexing.Tests
{
    public abstract class LauncherBase : Outputer
    {
        /// <summary>
        /// for column Индекс
        /// </summary>
        public const char IndicesSeparator = ',';

        private const string FilterString20 = "яш кя ля пс ую щб рю бж фг фю сг цх ак цз ск фм аю ип нв фп";
        private const string FilterString40 = FilterString20 + " ай ба не ду вы сю бв щд шы мч хе щг ти цт лх сх гя лс пь вз";
        private static readonly string[] Filter = FilterString40.Split(' ');
        
        /// <summary>
        /// XLS format is faster for reading
        /// </summary>
        public abstract string FileNameIn { get; }

        /// <summary>
        /// big files are slow in Xlsx for writing
        /// </summary>
        public virtual OutputType FileOutputType { get { return OutputType.Xlsx; } }

        public virtual bool ReadLegend { get { return false; } }

#region http://hist.tk/ory/EventsIndexing
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

        /// <summary>
        /// совпадения со сдвигом в 1-5 год
        /// </summary>
        [Test]
        public void DoWithShifts()
        {
            var man = new ExcelManager(Path.Combine(Folder, FileNameIn));
            var graphByYears = GroupEventsByYear(man);
            DoWithShift(graphByYears, 1);
            DoWithShift(graphByYears, 2);
            DoWithShift(graphByYears, 3);
            DoWithShift(graphByYears, 4);
            DoWithShift(graphByYears, 5);
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

        protected /*private*/ void DoWithShift(Graphes graphByYears, int shift)
        {
            var s = shift == 0 ? null : string.Format("_сдвиг_{0}год", shift);
            var rows = EventtypesByYears.Do(graphByYears, shift);
            var subname = string.Format(
                "сводная{0}{1}{2}",
                graphByYears.TwoAndMoreEventtypes ? "_filter2_" : null,
                graphByYears.Legends != null ? "Группирована" : null,
                s);
            Output(
                rows,
                subname,
                FileOutputType);

            // Output(rows, "проценты" + s, true, null, true);
            if (!graphByYears.TwoAndMoreEventtypes && graphByYears.Legends == null)
                Output(
                    EventtypesByYears.Do(graphByYears, shift), // 822, 1852),
                    "сводная" + Filter.Length + s, 
                    OutputType.Csv, 
                    Filter);
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
    }
}
#endif
