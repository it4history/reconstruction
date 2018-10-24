#if DEBUG

using System.Collections.Generic;
using System.IO;
using System.Text;
using Logy.Api.Mw.Excel;
using NUnit.Framework;

namespace Routines.Excel.EventsIndexing.Tests
{
    /// <summary>
    /// externally referenced, do not move or rename
    /// http://hist.tk/hw/Поиск_этапов_развития_цивилизации
    /// -5000	-2000	-1000	0	500	1000	1500	1800	1850	1900	1950
    /// 3000  2500  1700  1000  700 400 200 150 120 100
    /// </summary>
    public class StagesLauncher : FullDbLauncher
    {
        public override string FolderOut
        {
            get { return "outStages"; }
        }

        [Test]
        public void DoForGephi()
        {
            var eventsMan = new ExcelManager(Path.Combine(Folder, FileNameIn));
            var man = new ExcelManager(eventsMan, 1); // Легенда sheet
            var legends = GetLegend(man, "Подгруппа"); //"Поисковые слова");
            var byYears = GroupEventsByYear(eventsMan, legends);

            var exceptions = new List<string> { "примечания" };
            var sb = new StringBuilder();
            sb.AppendLine("Source,Target,Timeset");//",Weight,Type");
            foreach (var year in byYears.GetYears())
            {
                var nodes = byYears.GetNodes(year);
                foreach (var node in nodes)
                {
                    if (!exceptions.Contains(node))
                        foreach (var otherNode in nodes)
                        {
                            if (!exceptions.Contains(otherNode))
                            {
                                sb.AppendFormat(
                                    "{0},{1},<[{2}]>",
                                    node,
                                    otherNode,
                                    year);
                                sb.AppendLine();
                            }
                        }
                }
            }
            File.WriteAllText(GetOutputPath("поПодгруппам", OutputType.Csv), sb.ToString());
        }
    }
}
#endif