using System.Collections.Generic;
using System.Linq;

namespace Logy.Api.Mw.Excel
{
    public static class EventtypesByYears
    {
		/// <returns>graph; rowName, columnName, count</returns>
        public static Dictionary<string, Dictionary<string, int>> Do(
        	Dictionary<short, string> byYears,
        	string filter = null) 
		{
            var rows = new Dictionary<string, Dictionary<string, int>>();
            foreach (var indices in byYears.Values) {
	            var nodes = indices.Split(',').ToList()
	            	.Distinct()
	            	.Where(node => !string.IsNullOrEmpty(node) 
	            	       && (filter == null || filter.Split(' ').Contains(node)));
	            foreach(var node in nodes) {
        			if (!rows.ContainsKey(node))
        			{
        				var list = new Dictionary<string, int>();
        				rows.Add(node, list);
        			}
	            }
	            
                // make complete graph
	  			foreach (var node in nodes) {
		  			foreach (var node2 in nodes) {
                			var cols = rows[node];
                			if (!cols.ContainsKey(node2))
                				cols[node2] = 1;
                			else
                				cols[node2]++;
                	}
                }
            }

            return rows;
        }
    }
}