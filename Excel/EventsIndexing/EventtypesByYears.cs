using System;
using System.Collections.Generic;
using System.Linq;

namespace Routines.Excel.EventsIndexing
{
    public static class EventtypesByYears
    {
        /// <returns>graph; rowName, columnName, count</returns>
        public static Dictionary<string, Dictionary<string, int>> Do(
            Dictionary<int, string> byYears,
            string filter = null,
            int shift = 0,
            int? yearFrom = null,
            int? yearTo = null)
        {
            var rows = new Dictionary<string, Dictionary<string, int>>();

            foreach (var year in byYears.Keys
                .Where(year => (yearFrom == null || year >= yearFrom)
                               && (yearTo == null || year <= yearTo))
                .OrderBy(year => year))
            {
                // should be not null 
                var nodes = GetNodes(byYears, year, filter);
                foreach (var node in nodes)
                {
                    if (!rows.ContainsKey(node))
                    {
                        rows.Add(node, new Dictionary<string, int>());
                    }
                }

                var otherNodes = shift == 0
                    ? nodes // complete graph will be made
                    : GetNodes(byYears, year - shift, filter);

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

        private static IEnumerable<string> GetNodes(Dictionary<int, string> byYears, int year, string filter)
        {
            if (byYears.ContainsKey(year))
                return byYears[year].Split(',').ToList()
                    .Distinct()
                    .Where(node => !string.IsNullOrEmpty(node)
                                   && (filter == null || filter.Split(' ').Contains(node)));
            return null;
        }
    }
}