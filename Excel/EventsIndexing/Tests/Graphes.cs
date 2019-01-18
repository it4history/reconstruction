using System.Collections.Generic;
using System.Linq;

namespace Routines.Excel.EventsIndexing.Tests
{
    public class Graphes
    {
        public readonly Dictionary<int, string> ByYears = new Dictionary<int, string>();
        private readonly Dictionary<string, string> _legends;
        private readonly bool _twoAndMoreEventtypes;

        public Graphes(Dictionary<string, string> legends, bool twoAndMoreEventtypes)
        {
            _legends = legends;
            _twoAndMoreEventtypes = twoAndMoreEventtypes;
        }

        internal Graphes(Dictionary<int, string> graphByYears)
        {
            ByYears = graphByYears;
        }

        public bool TwoAndMoreEventtypes
        {
            get { return _twoAndMoreEventtypes; }
        }

        public Dictionary<string, string> Legends
        {
            get { return _legends; }
        }

        public void Add(int year, string indices)
        {
            ByYears[year] = (ByYears.ContainsKey(year) ? ByYears[year] : null) 
                + indices + LauncherBase.IndicesSeparator;
        }

        public IOrderedEnumerable<int> GetYears(int? yearFrom = null, int? yearTo = null)
        {
            return ByYears.Keys
                .Where(year => (yearFrom == null || year >= yearFrom)
                               && (yearTo == null || year <= yearTo))
                .OrderBy(year => year);
        }

        public IEnumerable<string> GetNodes(int? year, string[] filter = null)
        {
            if (year == null || !ByYears.ContainsKey(year.Value))
                return null;
            return ByYears[year.Value].Split(',').ToList()
                .Distinct()
                .Where(node => !string.IsNullOrEmpty(node)
                               && (filter == null || filter.Contains(node)));
        }
    }
}