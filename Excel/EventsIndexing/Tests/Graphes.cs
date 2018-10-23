using System.Collections.Generic;

namespace Routines.Excel.EventsIndexing.Tests
{
    public class Graphes
    {
        private readonly Dictionary<string, string> _legends;
        private readonly bool _twoAndMoreEventtypes;
        public readonly Dictionary<int, string> ByYears = new Dictionary<int, string>();

        public Graphes(Dictionary<string, string> legends, bool twoAndMoreEventtypes)
        {
            _legends = legends;
            _twoAndMoreEventtypes = twoAndMoreEventtypes;
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
    }
}