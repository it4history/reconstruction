using System.Collections.Generic;
using Routines.Excel.EventsIndexing.Tests;

namespace Routines.Excel.EventsIndexing
{
    /// <summary>
    /// externally referenced, do not move or rename
    /// </summary>
    public static class EventtypesByYears
    {
        /// <returns>graph; rowName, columnName, count</returns>
        public static Dictionary<string, Dictionary<string, int>> Do(
            Graphes byYears,
            // string[] filter = null, filtering make sense only during output
            int shift = 0,
            int? yearFrom = null,
            int? yearTo = null)
        {
            var rows = new Dictionary<string, Dictionary<string, int>>();

            foreach (var year in byYears.GetYears(yearFrom, yearTo))
            {
                var nodes = byYears.GetNodes(year);
                foreach (var node in nodes) // should be not null 
                {
                    if (!rows.ContainsKey(node))
                    {
                        rows.Add(node, new Dictionary<string, int>());
                    }
                }

                var otherNodes = shift == 0
                    ? nodes // complete graph will be made
                    : byYears.GetNodes(year + shift); //: byYears.GetNodes(previousYear, filter);

                if (otherNodes != null)
                    foreach (var node in nodes)
                    {
                        foreach (var node2 in otherNodes)
                        {
                            var cols = rows[node];
                            if (cols.ContainsKey(node2))
                                cols[node2]++;
                            else
                                cols[node2] = 1;
                        }
                    }
            }

            return rows;
        }
    }
}